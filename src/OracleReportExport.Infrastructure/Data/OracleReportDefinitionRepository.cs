using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Enums;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Interfaces;

namespace OracleReportExport.Infrastructure.Data
{
    public sealed class OracleReportDefinitionRepository : IReportDefinitionRepository
    {
        private readonly IOracleConnectionFactory _connectionFactory;
        private readonly string _connectionId;

        public OracleReportDefinitionRepository(
            IOracleConnectionFactory connectionFactory,
            string connectionId)
        {
            _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));
            _connectionId = connectionId ?? throw new ArgumentNullException(nameof(connectionId));
        }

        public async Task<IReadOnlyList<ReportDefinition>> GetAllAsync(CancellationToken ct = default)
        {
            using var conn = (OracleConnection)_connectionFactory.CreateConnection(_connectionId);
            await conn.OpenAsync(ct);

            var reportRows = await LoadReportsAsync(conn, ct);
            if (reportRows.Count == 0)
                return Array.Empty<ReportDefinition>();

            // Cargamos todo el resto de tablas (parametros y maestros)
            var mastersByReport = await LoadTableMastersAsync(conn, ct);
            var parametersByReport = await LoadParametersAsync(conn, ct);
           
            // Componemos el modelo de dominio final
            var result = new List<ReportDefinition>();

            foreach (var r in reportRows)
            {
                mastersByReport.TryGetValue(r.Id, out var masters);
                parametersByReport.TryGetValue(r.Id, out var parameters);

                var report = new ReportDefinition
                {
                    Id = r.Id,
                    Name = r.Name,
                    Category = r.Category,
                    Description = r.Description ?? string.Empty,
                    SourceType = r.SourceType,
                    SqlForStations = r.SqlForStations,
                    SqlForCentral = r.SqlForCentral,
                    SqlFileForStations = r.SqlFileForStations,
                    SqlFileForCentral = r.SqlFileForCentral,
                    TableMasterForParameters = masters ?? Array.Empty<TableMasterParameterDefinition>(),
                    Parameters = parameters ?? Array.Empty<ReportParameterDefinition>()
                };
                report.DefaultConnectionIds= await LoadConnectionIdsForReport(report.Id, conn, ct);            
                result.Add(report);
            }

            return result;
        }

        public async Task<ReportDefinition?> GetByIdAsync(string id, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Id no puede ser nulo o vacío.", nameof(id));

            var all = await GetAllAsync(ct);
            return all.FirstOrDefault(r => string.Equals(r.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        // ----- Carga de cabecera de informes -----

        private sealed class ReportRow
        {
            public string Id { get; init; } = string.Empty;
            public string Name { get; init; } = string.Empty;
            public string Category { get; init; } = string.Empty;
            public string? Description { get; init; }
            public ReportSourceType SourceType { get; init; }
            public string? SqlForStations { get; init; }
            public string? SqlForCentral { get; init; }
            public string? SqlFileForStations { get; init; }
            public string? SqlFileForCentral { get; init; }
        }

        private static async Task<List<ReportRow>> LoadReportsAsync(OracleConnection conn, CancellationToken ct)
        {
            const string sql = @"
                    SELECT
                        REPORT_ID,
                        NAME,
                        CATEGORY,
                        DESCRIPTION,
                        SOURCE_TYPE,
                        SQL_FOR_STATIONS,
                        SQL_FOR_CENTRAL,
                        SQL_FILE_FOR_STN,
                        SQL_FILE_FOR_CEN
                    FROM RPT_REPORT_DEFINITION
                    WHERE IS_ACTIVE = -1
                    ORDER BY CATEGORY, NAME";

            using var cmd = new OracleCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.Default, ct);

            var list = new List<ReportRow>();

            while (await reader.ReadAsync(ct))
            {
                var id = reader.GetString(0);
                var name = reader.GetString(1);
                var category = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
                var description = reader.IsDBNull(3) ? null : reader.GetString(3);
                var sourceTypeStr = reader.IsDBNull(4) ? "Estacion" : reader.GetString(4);

                if (!Enum.TryParse<ReportSourceType>(sourceTypeStr, true, out var sourceType))
                {
                    sourceType = ReportSourceType.Estacion;
                }

                string? sqlStations = reader.IsDBNull(5) ? null : reader.GetString(5);
                string? sqlCentral = reader.IsDBNull(6) ? null : reader.GetString(6);
                string? sqlFileStations = reader.IsDBNull(7) ? null : reader.GetString(7);
                string? sqlFileCentral = reader.IsDBNull(8) ? null : reader.GetString(8);

                list.Add(new ReportRow
                {
                    Id = id,
                    Name = name,
                    Category = category,
                    Description = description,
                    SourceType = sourceType,
                    SqlForStations = sqlStations,
                    SqlForCentral = sqlCentral,
                    SqlFileForStations = sqlFileStations,
                    SqlFileForCentral = sqlFileCentral
                });
            }

            return list;
        }

        // ----- Carga de tablas maestras (TableMasterForParameters) -----

        private static async Task<Dictionary<string, IReadOnlyList<TableMasterParameterDefinition>>> LoadTableMastersAsync(
            OracleConnection conn,
            CancellationToken ct)
        {
            const string sql = @"
                        SELECT
                            tm.TM_ID,
                            tm.REPORT_ID,
                            tm.NAME,
                            tm.LABEL,
                            tm.TYPE,
                            tm.IS_REQUIRED,
                            tm.ID_COLUMN,
                            tm.TEXT_COLUMN,
                            tm.SQL_QUERY_MASTER,
                            tr.REQUIRED_VALUE,
                            tr.ORDER_INDEX
                        FROM RPT_REPORT_TABLE_MASTER tm
                        LEFT JOIN RPT_REPORT_TM_REQUIRED_VALUE tr
                            ON tr.TM_ID = tm.TM_ID
                        ORDER BY tm.REPORT_ID, tm.TM_ID, tr.ORDER_INDEX";

            using var cmd = new OracleCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.Default, ct);

            var mastersByTmId = new Dictionary<int, (string ReportId, TableMasterParameterDefinition Def)>();

            while (await reader.ReadAsync(ct))
            {
                var tmId = reader.GetInt32(0);
                var reportId = reader.GetString(1);

                if (!mastersByTmId.TryGetValue(tmId, out var entry))
                {
                    var name = reader.IsDBNull(2) ? null : reader.GetString(2);
                    var label = reader.IsDBNull(3) ? null : reader.GetString(3);
                    var type = reader.IsDBNull(4) ? null : reader.GetString(4);
                    var isRequiredVal = reader.IsDBNull(5) ? 0 : reader.GetInt32(5);
                    var idColumn = reader.IsDBNull(6) ? null : reader.GetString(6);
                    var textColumn = reader.IsDBNull(7) ? null : reader.GetString(7);
                    var sqlQuery = reader.IsDBNull(8) ? null : reader.GetString(8);

                    var def = new TableMasterParameterDefinition
                    {
                        Name = name,
                        Label = label,
                        Type = type,
                        IsRequired = isRequiredVal != 0,
                        Id = idColumn,
                        Text = textColumn,
                        SqlQueryMaster = sqlQuery,
                        ValuesRequired = new List<string?>()
                    };

                    entry = (reportId, def);
                    mastersByTmId[tmId] = entry;
                }

                // REQUIRED_VALUE puede ser null si no hay filas en RPT_REPORT_TM_REQUIRED_VALUE
                if (!reader.IsDBNull(9))
                {
                    var value = reader.GetString(9);
                    if (entry.Def.ValuesRequired == null)
                        entry.Def.ValuesRequired = new List<string?>();
                    entry.Def.ValuesRequired.Add(value);
                }
            }

            // Agrupamos por REPORT_ID
            var result = new Dictionary<string, IReadOnlyList<TableMasterParameterDefinition>>();

            foreach (var group in mastersByTmId.Values.GroupBy(x => x.ReportId))
            {
                var list = group.Select(x => x.Def).ToList();
                result[group.Key] = list;
            }

            return result;
        }

        private static async Task<List<String>> LoadConnectionIdsForReport(String reportId,
            OracleConnection conn,
            CancellationToken ct)
        {
            List<string> result = new List<string>();
            const string sql = @"
                    SELECT CONNECTION_ID
                FROM RPT_REPORT_CONNECTION
                WHERE REPORT_ID = :REPORT_ID";

            using (var cmd = new OracleCommand(sql, conn)
            {
                BindByName = true
            })
            {

                cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = reportId;
                using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.Default, ct);
                while (await reader.ReadAsync(ct))
                    result.Add(reader.GetString(0));
            };
        
            return (List<String>)result;
        }
        // ----- Carga de parámetros (Parameters) -----

        private static async Task<Dictionary<string, IReadOnlyList<ReportParameterDefinition>>> LoadParametersAsync(
            OracleConnection conn,
            CancellationToken ct)
        {
            
            const string sql = @"
                    SELECT
                        p.PARAM_ID,
                        p.REPORT_ID,
                        p.NAME,
                        p.LABEL,
                        p.TYPE,
                        p.IS_REQUIRED,
                        p.ALLOWED_VALUES_JSON,
                        p.BUSQUEDA_LIKE,
                        v.KEY_INT,
                        v.VALUE_TEXT
                    FROM RPT_REPORT_PARAMETER p
                    LEFT JOIN RPT_REPORT_PARAM_VALUE v
                        ON v.PARAM_ID = p.PARAM_ID
                    ORDER BY p.REPORT_ID, p.PARAM_ID, v.KEY_INT";

            using var cmd = new OracleCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.Default, ct);

            var byReport = new Dictionary<string, List<ReportParameterDefinition>>();
            var byParamId = new Dictionary<int, ReportParameterDefinition>();

            while (await reader.ReadAsync(ct))
            {
                var paramId = reader.GetInt32(0);
                var reportId = reader.GetString(1);

                if (!byParamId.TryGetValue(paramId, out var param))
                {
                    var name = reader.GetString(2);
                    var label = reader.GetString(3);
                    var type = reader.GetString(4);
                    var isReqVal = reader.IsDBNull(5) ? 0 : reader.GetInt32(5);

                    string[]? allowedValues = null;
                    if (!reader.IsDBNull(6))
                    {
                        var json = reader.GetString(6);
                        if (!string.IsNullOrWhiteSpace(json))
                        {
                            try
                            {
                                allowedValues = JsonSerializer.Deserialize<string[]>(json);
                            }
                            catch
                            {
                                // Si falla el JSON, lo ignoramos
                                allowedValues = null;
                            }
                        }
                    }

                    bool? busquedaLike = null;
                    if (!reader.IsDBNull(7))
                    {
                        var likeVal = reader.GetInt32(7);
                        busquedaLike = likeVal != 0;
                    }

                    param = new ReportParameterDefinition
                    {
                        Name = name,
                        Label = label,
                        Type = type,
                        IsRequired = isReqVal != 0,
                        AllowedValues = allowedValues,
                        Values = new List<IntCodeItem>(),
                        BusquedaLike = busquedaLike
                    };

                    byParamId[paramId] = param;

                    if (!byReport.TryGetValue(reportId, out var list))
                    {
                        list = new List<ReportParameterDefinition>();
                        byReport[reportId] = list;
                    }
                    list.Add(param);
                }

                // Valores fijos del parámetro (IntCodeItem)
                if (!reader.IsDBNull(8))
                {
                    var key = reader.GetInt32(8);
                    var val = reader.IsDBNull(9) ? null : reader.GetString(9);
                    param.Values.Add(new IntCodeItem
                    {
                        Key = key,
                        Value = val
                    });
                }
            }

            // Convertimos a IReadOnlyList
            var result = byReport.ToDictionary(
                kvp => kvp.Key,
                kvp => (IReadOnlyList<ReportParameterDefinition>)kvp.Value);

            return result;
        }

        ////public async Task SaveAsync(ReportDefinition report, CancellationToken ct = default)
        ////{
        ////    if (report == null)
        ////        throw new ArgumentNullException(nameof(report));
        ////    try
        ////    {

        ////        using var conn = (OracleConnection)_connectionFactory.CreateConnection(_connectionId);
        ////        await conn.OpenAsync(ct);


        ////        using var tx = conn.BeginTransaction();

        ////        const string sql = @"
        ////        INSERT INTO RPT_REPORT_DEFINITION
        ////            (REPORT_ID, NAME, CATEGORY, DESCRIPTION, SOURCE_TYPE,
        ////             SQL_FOR_STATIONS, SQL_FOR_CENTRAL, IS_ACTIVE)
        ////        VALUES
        ////            (:REPORT_ID, :NAME, :CATEGORY, :DESCRIPTION, :SOURCE_TYPE,
        ////             :SQL_FOR_STATIONS, :SQL_FOR_CENTRAL, -1)";
        ////        using var cmd = new OracleCommand(sql, conn)
        ////        {
        ////            BindByName = true 
        ////        };
        ////        cmd.Transaction=tx;
        ////        cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = report.Id;
        ////        cmd.Parameters.Add("NAME", OracleDbType.Varchar2).Value = report.Name;
        ////        cmd.Parameters.Add("CATEGORY", OracleDbType.Varchar2).Value = report.Category;
        ////        cmd.Parameters.Add("DESCRIPTION", OracleDbType.Varchar2).Value =
        ////            (object?)report.Description ?? DBNull.Value;
        ////        cmd.Parameters.Add("SOURCE_TYPE", OracleDbType.Varchar2).Value = report.SourceType.ToString();
        ////        cmd.Parameters.Add("SQL_FOR_STATIONS", OracleDbType.Clob).Value =
        ////            (object?)report.SqlForStations ?? DBNull.Value;
        ////        cmd.Parameters.Add("SQL_FOR_CENTRAL", OracleDbType.Clob).Value =
        ////            (object?)report.SqlForCentral ?? DBNull.Value;

        ////        using var registration = ct.Register(() => cmd.Cancel());

        ////        await cmd.ExecuteNonQueryAsync(ct);

        ////        tx.Commit();
        ////    }
        ////    catch (Exception ex)
        ////    {

        ////        throw ex;
        ////    }
        ////}
        ///

        public async Task SaveAsync(ReportDefinition report, CancellationToken ct = default)
        {
            if (report == null)
                throw new ArgumentNullException(nameof(report));

            using var conn = (OracleConnection)_connectionFactory.CreateConnection(_connectionId);
            await conn.OpenAsync(ct);

            using var tx = conn.BeginTransaction();

            try
            {
                // --------------------------------------------------------------------
                // 1) INSERT PRINCIPAL: RPT_REPORT_DEFINITION
                // --------------------------------------------------------------------
                const string sqlReport = @"
                                    INSERT INTO RPT_REPORT_DEFINITION
                                        (REPORT_ID, NAME, CATEGORY, DESCRIPTION, SOURCE_TYPE,
                                         SQL_FOR_STATIONS, SQL_FOR_CENTRAL, IS_ACTIVE)
                                    VALUES
                                        (:REPORT_ID, :NAME, :CATEGORY, :DESCRIPTION, :SOURCE_TYPE,
                                         :SQL_FOR_STATIONS, :SQL_FOR_CENTRAL, -1)";

                using (var cmd = new OracleCommand(sqlReport, conn)
                {
                    BindByName = true
                })
                {
                    cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = report.Id;
                    cmd.Parameters.Add("NAME", OracleDbType.Varchar2).Value = report.Name;
                    cmd.Parameters.Add("CATEGORY", OracleDbType.Varchar2).Value = report.Category;
                    cmd.Parameters.Add("DESCRIPTION", OracleDbType.Varchar2).Value =
                        (object?)report.Description ?? DBNull.Value;
                    cmd.Parameters.Add("SOURCE_TYPE", OracleDbType.Varchar2).Value = report.SourceType.ToString();
                    cmd.Parameters.Add("SQL_FOR_STATIONS", OracleDbType.Clob).Value =
                        (object?)report.SqlForStations ?? DBNull.Value;
                    cmd.Parameters.Add("SQL_FOR_CENTRAL", OracleDbType.Clob).Value =
                        (object?)report.SqlForCentral ?? DBNull.Value;

                    using var reg = ct.Register(() => cmd.Cancel());
                    await cmd.ExecuteNonQueryAsync(ct);
                }

                // --------------------------------------------------------------------
                // 1.1) INSERT RPT_REPORT_CONNECTION Tabla intermedia
                // --------------------------------------------------------------------
                // Después del INSERT en RPT_REPORT_DEFINITION
                if (report.DefaultConnectionIds != null && report.DefaultConnectionIds.Count > 0)
                {
                    const string sqlConn = @"
                        INSERT INTO RPT_REPORT_CONNECTION (REPORT_ID, CONNECTION_ID)
                        VALUES (:REPORT_ID, :CONNECTION_ID)";

                    using var cmdConn = new OracleCommand(sqlConn, conn)
                    {
                        BindByName = true
                    };
                    cmdConn.Parameters.Add("REPORT_ID", OracleDbType.Varchar2);
                    cmdConn.Parameters.Add("CONNECTION_ID", OracleDbType.Varchar2);

                    foreach (var connId in report.DefaultConnectionIds.Distinct())
                    {
                        cmdConn.Parameters["REPORT_ID"].Value = report.Id;
                        cmdConn.Parameters["CONNECTION_ID"].Value = connId;
                        await cmdConn.ExecuteNonQueryAsync(ct);
                    }
                }




                // --------------------------------------------------------------------
                // 2) INSERT RPT_REPORT_PARAMETER + RPT_REPORT_PARAM_VALUE
                // --------------------------------------------------------------------
                if (report.Parameters != null && report.Parameters.Count > 0)
                {
                    const string sqlParam = @"
                                INSERT INTO RPT_REPORT_PARAMETER
                                    (PARAM_ID, REPORT_ID, NAME, LABEL, TYPE,
                                     IS_REQUIRED, ALLOWED_VALUES_JSON, BUSQUEDA_LIKE)
                                VALUES
                                    (:PARAM_ID, :REPORT_ID, :NAME, :LABEL, :TYPE,
                                     :IS_REQUIRED, :ALLOWED_VALUES_JSON, :BUSQUEDA_LIKE)";

                    const string sqlParamValue = @"
                                INSERT INTO RPT_REPORT_PARAM_VALUE
                                    (PARAM_VALUE_ID, PARAM_ID, KEY_INT, VALUE_TEXT)
                                VALUES
                                    (:PARAM_VALUE_ID, :PARAM_ID, :KEY_INT, :VALUE_TEXT)";

                    foreach (var p in report.Parameters)
                    {
                        // Obtener nuevo PARAM_ID desde la secuencia
                        var paramId = await GetNextSequenceValueAsync("RPT_PARAM_SEQ", conn, tx, ct);

                        // 2.1 Fila en RPT_REPORT_PARAMETER
                        using (var cmdParam = new OracleCommand(sqlParam, conn)
                        {
                            BindByName = true
                        })
                        {
                            cmdParam.Parameters.Add("PARAM_ID", OracleDbType.Int64).Value = paramId;
                            cmdParam.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = report.Id;
                            cmdParam.Parameters.Add("NAME", OracleDbType.Varchar2).Value = p.Name;
                            cmdParam.Parameters.Add("LABEL", OracleDbType.Varchar2).Value =
                                (object?)p.Label ?? DBNull.Value;
                            cmdParam.Parameters.Add("TYPE", OracleDbType.Varchar2).Value = p.Type ?? "text";
                            cmdParam.Parameters.Add("IS_REQUIRED", OracleDbType.Int16).Value = p.IsRequired ? -1 : 0;

                            // De momento dejamos el JSON a NULL (para AdHoc normalmente no lo necesitas)
                            cmdParam.Parameters.Add("ALLOWED_VALUES_JSON", OracleDbType.Clob).Value = DBNull.Value;

                            if (p.BusquedaLike.HasValue)
                                cmdParam.Parameters.Add("BUSQUEDA_LIKE", OracleDbType.Int16).Value =
                                    p.BusquedaLike.Value ? -1 : 0;
                            else
                                cmdParam.Parameters.Add("BUSQUEDA_LIKE", OracleDbType.Int16).Value = DBNull.Value;

                            using var regParam = ct.Register(() => cmdParam.Cancel());
                            await cmdParam.ExecuteNonQueryAsync(ct);
                        }

                        // 2.2 Valores de parámetro (lista p.Values -> RPT_REPORT_PARAM_VALUE)
                        if (p.Values != null && p.Values.Count > 0)
                        {
                            foreach (var v in p.Values)
                            {
                                var paramValueId = await GetNextSequenceValueAsync("RPT_PARAM_VALUE_SEQ", conn, tx, ct);

                                using var cmdVal = new OracleCommand(sqlParamValue, conn)
                                {
                                    BindByName = true
                                };

                                cmdVal.Parameters.Add("PARAM_VALUE_ID", OracleDbType.Int64).Value = paramValueId;
                                cmdVal.Parameters.Add("PARAM_ID", OracleDbType.Int64).Value = paramId;
                                cmdVal.Parameters.Add("KEY_INT", OracleDbType.Int32).Value = v.Key;
                                cmdVal.Parameters.Add("VALUE_TEXT", OracleDbType.Varchar2).Value =
                                    (object?)v.Value ?? DBNull.Value;

                                using var regVal = ct.Register(() => cmdVal.Cancel());
                                await cmdVal.ExecuteNonQueryAsync(ct);
                            }
                        }
                    }
                }

                // --------------------------------------------------------------------
                // 3) INSERT RPT_REPORT_TABLE_MASTER + RPT_REPORT_TM_REQUIRED_VALUE
                // --------------------------------------------------------------------
                if (report.TableMasterForParameters != null &&
                    report.TableMasterForParameters.Count > 0)
                {
                    const string sqlTm = @"
                            INSERT INTO RPT_REPORT_TABLE_MASTER
                                (TM_ID, REPORT_ID, NAME, LABEL, TYPE, IS_REQUIRED,
                                 ID_COLUMN, TEXT_COLUMN, SQL_QUERY_MASTER)
                            VALUES
                                (:TM_ID, :REPORT_ID, :NAME, :LABEL, :TYPE, :IS_REQUIRED,
                                 :ID_COLUMN, :TEXT_COLUMN, :SQL_QUERY_MASTER)";

                    const string sqlTmReq = @"
                        INSERT INTO RPT_REPORT_TM_REQUIRED_VALUE
                            (TM_ID, TM_REQ_ID, REQUIRED_VALUE, ORDER_INDEX)
                        VALUES
                            (:TM_ID, :TM_REQ_ID, :REQUIRED_VALUE, :ORDER_INDEX)";

                    foreach (var tm in report.TableMasterForParameters)
                    {
                        var tmId = await GetNextSequenceValueAsync("RPT_TM_SEQ", conn, tx, ct);

                        // 3.1 Fila en RPT_REPORT_TABLE_MASTER
                        using (var cmdTm = new OracleCommand(sqlTm, conn)
                        {
                            BindByName = true
                        })
                        {
                            cmdTm.Parameters.Add("TM_ID", OracleDbType.Int64).Value = tmId;
                            cmdTm.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = report.Id;
                            cmdTm.Parameters.Add("NAME", OracleDbType.Varchar2).Value = tm.Name;
                            cmdTm.Parameters.Add("LABEL", OracleDbType.Varchar2).Value =
                                (object?)tm.Label ?? DBNull.Value;
                            cmdTm.Parameters.Add("TYPE", OracleDbType.Varchar2).Value = tm.Type ?? "combobox";
                            cmdTm.Parameters.Add("IS_REQUIRED", OracleDbType.Int16).Value = tm.IsRequired ? -1 : 0;
                            cmdTm.Parameters.Add("ID_COLUMN", OracleDbType.Varchar2).Value =
                                (object?)tm.Id ?? DBNull.Value;
                            cmdTm.Parameters.Add("TEXT_COLUMN", OracleDbType.Varchar2).Value =
                                (object?)tm.Text ?? DBNull.Value;
                            cmdTm.Parameters.Add("SQL_QUERY_MASTER", OracleDbType.Clob).Value =
                                (object?)tm.SqlQueryMaster ?? DBNull.Value;

                            using var regTm = ct.Register(() => cmdTm.Cancel());
                            await cmdTm.ExecuteNonQueryAsync(ct);
                        }

                        // 3.2 Valores requeridos (RPT_REPORT_TM_REQUIRED_VALUE)
                        if (tm.ValuesRequired != null && tm.ValuesRequired.Count > 0)
                        {
                            int orderIndex = 1;

                            foreach (var required in tm.ValuesRequired)
                            {
                                var tmReqId = await GetNextSequenceValueAsync("RPT_TM_REQUIRED_SEQ", conn, tx, ct);

                                using var cmdReq = new OracleCommand(sqlTmReq, conn)
                                {
                                    BindByName = true
                                };

                                cmdReq.Parameters.Add("TM_ID", OracleDbType.Int64).Value = tmId;
                                cmdReq.Parameters.Add("TM_REQ_ID", OracleDbType.Int64).Value = tmReqId;
                                cmdReq.Parameters.Add("REQUIRED_VALUE", OracleDbType.Varchar2).Value = required;
                                cmdReq.Parameters.Add("ORDER_INDEX", OracleDbType.Int32).Value = orderIndex++;

                                using var regReq = ct.Register(() => cmdReq.Cancel());
                                await cmdReq.ExecuteNonQueryAsync(ct);
                            }
                        }
                    }
                }

                // --------------------------------------------------------------------
                // 4) COMMIT FINAL
                // --------------------------------------------------------------------
                tx.Commit();
            }
            catch
            {
                try { tx.Rollback(); } catch { /* ignore */ }
                throw; // re-lanza manteniendo el stack original
            }
        }

        /// <summary>
        /// Devuelve NEXTVAL de una secuencia dentro de la misma conexión+transacción.
        /// </summary>
        private static async Task<long> GetNextSequenceValueAsync(
            string sequenceName,
            OracleConnection conn,
            OracleTransaction tx,
            CancellationToken ct)
        {
            var sql = $"SELECT {sequenceName}.NEXTVAL FROM DUAL";
            using var cmd = new OracleCommand(sql, conn);
            var result = await cmd.ExecuteScalarAsync(ct);
            return Convert.ToInt64(result);
        }


    }
}
