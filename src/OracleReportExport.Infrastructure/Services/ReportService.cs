using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Class;
using OracleReportExport.Domain.Enums;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Interfaces;
using Serilog;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OracleReportExport.Infrastructure.Services
{
    public sealed class ReportService : IReportService
    {
        private readonly IReportDefinitionRepository _definitions;
        private readonly IQueryExecutor _queryExecutor;

        public ReportService(
            IReportDefinitionRepository definitions,
            IQueryExecutor queryExecutor)
        {
            _definitions = definitions;
            _queryExecutor = queryExecutor;
        }

        public Task<IReadOnlyList<ReportDefinition>> GetAvailableReportsAsync(
            CancellationToken ct = default)
            => _definitions.GetAllAsync(ct);


        public async Task<ReportQueryResult> ExecuteReportAsync(ReportDefinition report,
                        IReadOnlyDictionary<string, object?> parameterValues,
                       List<ConnectionInfo> targetConnection,
                        CancellationToken ct = default)
        {
            if (report == null)
                throw new ArgumentNullException(nameof(report));

            var def = await _definitions.GetByIdAsync(report.Id);

            if (def is null)
                throw new InvalidOperationException(
                    $"No se encontró la definición del informe '{report.Id}'.");

            if (string.IsNullOrWhiteSpace(report.SqlFileForStations) &&
                string.IsNullOrWhiteSpace(report.SqlForStations) &&
                string.IsNullOrEmpty(report.SqlForCentral) &&
                string.IsNullOrEmpty(report.SqlForStations))
                throw new InvalidOperationException("El informe no tiene SQL configurada.");

            string sql=String.Empty;
            switch (report.SourceType)
            {
                case OracleReportExport.Domain.Enums.ReportSourceType.Central:
                    sql = report.SqlForCentral ?? string.Empty;
                    break;
                case OracleReportExport.Domain.Enums.ReportSourceType.Estacion:
                    sql = report.SqlForStations ?? string.Empty;
                    break;
            }
                //  Sustituir tokens del txt de  {TiposVehiculoList} y {CategoriasList}
                if (parameterValues.TryGetValue("TiposVehiculoList", out var tvListObj) &&
                    tvListObj is string tvList)
                {
                    sql = sql.Replace("{TiposVehiculoList}", tvList);
                }

                if (parameterValues.TryGetValue("CategoriasList", out var catListObj) &&
                    catListObj is string catList)
                {
                    sql = sql.Replace("{CategoriasList}", catList);
                }

            DataTable? combined = null;
            var timeoutConnections = new List<string>();
            foreach (var connectionId in targetConnection.ToList())
            {
                try
                {
                    var table = await _queryExecutor.ExecuteQueryAsync(
                                sql,
                                parameterValues,
                                connectionId,
                                report.Id,
                                ct);

                    if (combined is null)
                    {
                        combined = table.Clone();
                        combined.Columns.Add("CONEXION_ESTACION", typeof(string));
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        var newRow = combined.NewRow();

                        foreach (DataColumn col in table.Columns)
                        {
                            newRow[col.ColumnName] = row[col];
                        }

                        newRow["CONEXION_ESTACION"] = connectionId.ToString();
                        combined.Rows.Add(newRow);
                    }
                }
                catch (OracleException ex) when (ex.Number == 50000) //Timeout
                {
                    Log.Warning(ex,
                        "Timeout en la conexión {ConnectionId} para el informe {ReportId}",
                        connectionId, report.Id);
                    timeoutConnections.Add(connectionId.ToString());
                }
            }

            return new ReportQueryResult
            {
                Data = combined ?? new DataTable(),
                TimeoutConnections = timeoutConnections
            };

        }


        public async Task<ReportQueryResult> ExecuteNonQueryAsync(string? sql, IReadOnlyDictionary<string, object?>? parameterValues, List<ConnectionInfo> targetConnection, CancellationToken ct = default)
        {
            DataTable? combined = null;
            int resultNonQuery = 0;
            var timeoutConnections = new List<string>();
            if (string.IsNullOrWhiteSpace(sql))
                throw new ArgumentNullException(nameof(sql), "La consulta SQL no puede estar vacía.");

            var kind = ClassSql.ClassifySql(sql);
            int totalAffectedRows = 0;

            foreach (var connectionId in targetConnection.ToList())
            {
                try
                {
                    resultNonQuery += await _queryExecutor.ExecuteNonQueryAsync(
                                sql,
                                parameterValues,
                                connectionId,
                                String.Empty,
                                ct);
                    if(kind==SqlKind.DdlSafe || kind==SqlKind.DdlDangerous || kind==SqlKind.PlSqlBlock )
                    {
                        //si es un alter sumo uno
                        //porque no devuelve filas afectadas las sentencias DDL
                        if (resultNonQuery == -1)
                        {
                            resultNonQuery = 1;
                            totalAffectedRows+= resultNonQuery;
                        }
                        else
                            totalAffectedRows += 1;
                    }
                }
                catch (OracleException ex) when (ex.Number == 50000) //Timeout
                {
                    Log.Warning(ex,
                        "Timeout en la conexión {ConnectionId} para la consulta ejecutada");
                    timeoutConnections.Add(connectionId.ToString());
                }
            }

            return new ReportQueryResult
            {
                Data = null,
                TimeoutConnections = timeoutConnections,
                RowsAffected = totalAffectedRows,
                Kind = kind
            };
        }


        public async Task<ReportQueryResult> ExecuteSQLAdHocAsync(string? sql, IReadOnlyDictionary<string, object?> ?parameterValues, List<ConnectionInfo> targetConnection, CancellationToken ct = default)
        {
            DataTable? combined = null;
            var timeoutConnections = new List<string>();
            if(string.IsNullOrWhiteSpace(sql))
                throw new ArgumentNullException(nameof(sql), "La consulta SQL no puede estar vacía.");
            foreach (var connectionId in targetConnection.ToList())
            {
                try
                {
                    var table = await _queryExecutor.ExecuteQueryAsync(
                                sql,
                                parameterValues,
                                connectionId,
                                String.Empty,
                                ct);

                    if (combined is null)
                    {
                        combined = table.Clone();
                        combined.Columns.Add("CONEXION_ESTACION", typeof(string));
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        var newRow = combined.NewRow();

                        foreach (DataColumn col in table.Columns)
                        {
                            newRow[col.ColumnName] = row[col];
                        }

                        newRow["CONEXION_ESTACION"] = connectionId.ToString()?.Trim()??"";
                        combined.Rows.Add(newRow);
                    }
                }
                catch (OracleException ex) when (ex.Number == 50000) //Timeout
                {
                    Log.Warning(ex,
                        "Timeout en la conexión {ConnectionId} para la consulta ejecutada");
                    timeoutConnections.Add(connectionId.ToString());
                }
            }

            return new ReportQueryResult
            {
                Data = combined ?? new DataTable(),
                TimeoutConnections = timeoutConnections
            };
        }

        public  async Task<bool> ValidateSqlSyntaxAsync(string sql, ConnectionInfo connection, CancellationToken ct)
        {
            return await _queryExecutor.ValidateSqlSyntaxAsync(
                sql,
                connection,
                ct);
        }

        public async Task SaveAsync(ReportDefinition report, CancellationToken ct = default)
        {
            try
            {
                if(report == null)
                    throw new ArgumentNullException(nameof(report));
                    await _definitions.SaveAsync(report, ct);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

      
    }
}

