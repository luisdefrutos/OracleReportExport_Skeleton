using System;
using System.Data;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Configuration;

namespace OracleReportExport.Tools.Migration
{
    internal static class Program
    {
        static async Task<int> Main(string[] args)
        {
            try
            {
                // 1) Ruta del Reports.json
                // Si pasas una ruta por argumento, la usa; si no, usa la carpeta Configuration del ejecutable.
                var baseDir = AppContext.BaseDirectory;
                var defaultReportsPath = Path.Combine(baseDir, "Configuration", "Reports.json");
                var reportsPath = args.Length > 0 ? args[0] : defaultReportsPath;
                var reportsDir = Path.GetDirectoryName(reportsPath)
                 ?? baseDir;

                if (!File.Exists(reportsPath))
                {
                    Console.WriteLine($"No se encuentra Reports.json en: {reportsPath}");
                    return 1;
                }

                Console.WriteLine($"Leyendo definiciones desde: {reportsPath}");

                // 2) Connection string
                // Opción simple: la pones aquí a mano o la pasas como segundo argumento.
                var connectionString = args.Length > 1
                    ? args[1]
                    : "Data Source=(DESCRIPTION =(SDU=65535)(ADDRESS = (PROTOCOL = TCP)(HOST = 10.108.206.7)(PORT = 1521))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = ORCLDV.snprivate.vcnatisae.oraclevcn.com)));User Id=ITV;Password=ItvitV.-20"; // <-- RELLENAR

                // 3) Deserializar Reports.json usando tu modelo real
                var json = await File.ReadAllTextAsync(reportsPath);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                options.Converters.Add(new JsonStringEnumConverter());

                var root = JsonSerializer.Deserialize<ReportDefinitionRoot>(json, options);

                if (root == null || root.Reports == null || root.Reports.Count == 0)
                {
                    Console.WriteLine("No se han encontrado informes en Reports.json");
                    return 1;
                }

                Console.WriteLine($"Informes encontrados: {root.Reports.Count}");

                // 4) Conexión a Oracle y transacción
                using var connection = new OracleConnection(connectionString);
                await connection.OpenAsync();

                using var transaction = connection.BeginTransaction();

                foreach (var report in root.Reports)
                {
                    Console.WriteLine($"Insertando informe {report.Id} - {report.Name} ...");

                    var reportId = report.Id; // es string, coincide con REPORT_ID

                    await InsertReportAsync(connection, transaction, report, reportsDir);
                    await InsertTableMastersAsync(connection, transaction, reportId, report);
                    await InsertParametersAsync(connection, transaction, reportId, report);

                    Console.WriteLine($"Informe {report.Id} insertado correctamente.");
                }

                transaction.Commit();
                Console.WriteLine("Migración completada con éxito.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante la migración:");
                Console.WriteLine(ex);
                return -1;
            }
        }

        // ------------- INSERT INFORME ------------------

        private static async Task InsertReportAsync(
            OracleConnection connection,
            OracleTransaction transaction,
            ReportDefinition report,
            string reportsDir)
        {
            // 1) Resolver SQL para estaciones
            string? sqlStations = report.SqlForStations;

            if (string.IsNullOrWhiteSpace(sqlStations) &&
                !string.IsNullOrWhiteSpace(report.SqlFileForStations))
            {
                report.SqlForStations = TryReadSqlFile(reportsDir, report.SqlFileForStations);
            }

            // 2) Resolver SQL para central
            string? sqlCentral = report.SqlForCentral;

            if (string.IsNullOrWhiteSpace(sqlCentral) &&
                !string.IsNullOrWhiteSpace(report.SqlFileForCentral))
            {
                report.SqlForCentral = TryReadSqlFile(reportsDir, report.SqlFileForCentral);
            }


            try
            {
                const string sql = @"
                        INSERT INTO RPT_REPORT_DEFINITION
                            (REPORT_ID, NAME, CATEGORY, DESCRIPTION, SOURCE_TYPE,
                             SQL_FOR_STATIONS, SQL_FOR_CENTRAL,
                             SQL_FILE_FOR_STN, SQL_FILE_FOR_CEN, IS_ACTIVE)
                        VALUES
                            (:REPORT_ID, :NAME, :CATEGORY, :DESCRIPTION, :SOURCE_TYPE,
                             :SQL_FOR_STATIONS, :SQL_FOR_CENTRAL,
                             :SQL_FILE_FOR_STN, :SQL_FILE_FOR_CEN, :IS_ACTIVE)";

                using var cmd = new OracleCommand(sql)
                {
                    Connection = connection,
                    Transaction = transaction
                };

                cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = report.Id;
                cmd.Parameters.Add("NAME", OracleDbType.Varchar2).Value = report.Name;
                cmd.Parameters.Add("CATEGORY", OracleDbType.Varchar2).Value = report.Category;
                cmd.Parameters.Add("DESCRIPTION", OracleDbType.Varchar2).Value =
                    (object?)report.Description ?? DBNull.Value;
                cmd.Parameters.Add("SOURCE_TYPE", OracleDbType.Varchar2).Value =
                    report.SourceType.ToString(); // "Estacion" / "Central"

                cmd.Parameters.Add("SQL_FOR_STATIONS", OracleDbType.Clob).Value =
                    (object?)report.SqlForStations ?? DBNull.Value;
                cmd.Parameters.Add("SQL_FOR_CENTRAL", OracleDbType.Clob).Value =
                    (object?)report.SqlForCentral ?? DBNull.Value;

                cmd.Parameters.Add("SQL_FILE_FOR_STN", OracleDbType.Varchar2).Value =
                    (object?)report.SqlFileForStations ?? DBNull.Value;
                cmd.Parameters.Add("SQL_FILE_FOR_CEN", OracleDbType.Varchar2).Value =
                    (object?)report.SqlFileForCentral ?? DBNull.Value;

                // -1 = activo, 0 = inactivo
                cmd.Parameters.Add("IS_ACTIVE", OracleDbType.Int16).Value = -1;

                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante InsertReportAsync:");
                Console.WriteLine(ex);
                 
            }
        }


        private static string? TryReadSqlFile(string reportsDir, string sqlFileRelativePath)
        {
            try
            {
                // Si el path es absoluto, lo usamos tal cual; si es relativo, lo combinamos
                string fullPath;

                if (Path.IsPathRooted(sqlFileRelativePath))
                {
                    fullPath = sqlFileRelativePath;
                }
                else
                {
                    fullPath = Path.Combine(reportsDir, sqlFileRelativePath);
                }

                if (!File.Exists(fullPath))
                {
                    Console.WriteLine($"[AVISO] No se encuentra el fichero SQL: {fullPath}");
                    return null;
                }

                Console.WriteLine($"Leyendo SQL desde fichero: {fullPath}");
                return File.ReadAllText(fullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AVISO] Error leyendo fichero SQL '{sqlFileRelativePath}': {ex.Message}");
                return null;
            }
        }

        // ------------- INSERT TABLE MASTERS ------------------

        private static async Task InsertTableMastersAsync(
            OracleConnection connection,
            OracleTransaction transaction,
            string reportId,
            ReportDefinition report)
        {
            try
            {
                if (report.TableMasterForParameters == null)
                    return;

                foreach (var tm in report.TableMasterForParameters)
                {
                    if (tm == null)
                        continue;

                    // Insertamos el maestro y recuperamos TM_ID con RETURNING
                    const string sql = @"
                            INSERT INTO RPT_REPORT_TABLE_MASTER
                                (TM_ID, REPORT_ID, NAME, LABEL, TYPE, IS_REQUIRED,
                                 ID_COLUMN, TEXT_COLUMN, SQL_QUERY_MASTER)
                            VALUES
                                (RPT_TM_SEQ.NEXTVAL, :REPORT_ID, :NAME, :LABEL, :TYPE, :IS_REQUIRED,
                                 :ID_COLUMN, :TEXT_COLUMN, :SQL_QUERY_MASTER)
                            RETURNING TM_ID INTO :OUT_TM_ID";

                    using var cmd = new OracleCommand(sql)
                    {
                        Connection = connection,
                        Transaction = transaction
                    };

                    cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = reportId;
                    cmd.Parameters.Add("NAME", OracleDbType.Varchar2).Value = tm.Name ?? string.Empty;
                    cmd.Parameters.Add("LABEL", OracleDbType.Varchar2).Value = tm.Label ?? string.Empty;
                    cmd.Parameters.Add("TYPE", OracleDbType.Varchar2).Value = tm.Type ?? "ComboBox";

                    // -1 / 0 para IsRequired
                    cmd.Parameters.Add("IS_REQUIRED", OracleDbType.Int16).Value = tm.IsRequired ? -1 : 0;

                    cmd.Parameters.Add("ID_COLUMN", OracleDbType.Varchar2).Value = tm.Id ?? string.Empty;
                    cmd.Parameters.Add("TEXT_COLUMN", OracleDbType.Varchar2).Value = tm.Text ?? string.Empty;

                    cmd.Parameters.Add("SQL_QUERY_MASTER", OracleDbType.Clob).Value =
                        (object?)tm.SqlQueryMaster ?? DBNull.Value;

                    var outParam = new OracleParameter("OUT_TM_ID", OracleDbType.Int32)
                    {
                        Direction = ParameterDirection.Output
                    };
                    cmd.Parameters.Add(outParam);

                    var resultExecute = await cmd.ExecuteNonQueryAsync();
                    var tmId = Convert.ToInt32(outParam.Value.ToString());

                    // Ahora insertamos los ValuesRequired (si hay)
                    if (tm.ValuesRequired != null)
                    {
                        int index = 1;
                        foreach (var value in tm.ValuesRequired)
                        {
                            if (value == null)
                            {
                                index++;
                                continue;
                            }

                            await InsertTableMasterRequiredValueAsync(connection, transaction, tmId, value, index);
                            index++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante InsertTableMastersAsync:");
                Console.WriteLine(ex);

            }
        }

        private static async Task InsertTableMasterRequiredValueAsync(
            OracleConnection connection,
            OracleTransaction transaction,
            int tmId,
            string requiredValue,
            int orderIndex)
        {
            try
            {
                const string sql = @"
                    INSERT INTO RPT_REPORT_TM_REQUIRED_VALUE
                        (TM_REQ_ID, TM_ID, REQUIRED_VALUE, ORDER_INDEX)
                    VALUES
                        (RPT_TM_REQ_SEQ.NEXTVAL, :TM_ID, :REQUIRED_VALUE, :ORDER_INDEX)";

                using var cmd = new OracleCommand(sql)
                {
                    Connection = connection,
                    Transaction = transaction
                };

                cmd.Parameters.Add("TM_ID", OracleDbType.Int32).Value = tmId;
                cmd.Parameters.Add("REQUIRED_VALUE", OracleDbType.Varchar2).Value = requiredValue;
                cmd.Parameters.Add("ORDER_INDEX", OracleDbType.Int32).Value = orderIndex;

                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante InsertTableMasterRequiredValueAsync:");
                Console.WriteLine(ex);

            }
        }

        // ------------- INSERT PARÁMETROS ------------------

        private static async Task InsertParametersAsync(
            OracleConnection connection,
            OracleTransaction transaction,
            string reportId,
            ReportDefinition report)
        {
            try
            {
                if (report.Parameters == null)
                    return;

                foreach (var param in report.Parameters)
                {
                    const string sql = @"
                    INSERT INTO RPT_REPORT_PARAMETER
                        (PARAM_ID, REPORT_ID, NAME, LABEL, TYPE, IS_REQUIRED,
                         ALLOWED_VALUES_JSON, BUSQUEDA_LIKE)
                    VALUES
                        (RPT_PARAM_SEQ.NEXTVAL, :REPORT_ID, :NAME, :LABEL, :TYPE, :IS_REQUIRED,
                         :ALLOWED_VALUES_JSON, :BUSQUEDA_LIKE)
                    RETURNING PARAM_ID INTO :OUT_PARAM_ID";

                    using var cmd = new OracleCommand(sql)
                    {
                        Connection = connection,
                        Transaction = transaction
                    };

                    cmd.Parameters.Add("REPORT_ID", OracleDbType.Varchar2).Value = reportId;
                    cmd.Parameters.Add("NAME", OracleDbType.Varchar2).Value = param.Name;
                    cmd.Parameters.Add("LABEL", OracleDbType.Varchar2).Value = param.Label;
                    cmd.Parameters.Add("TYPE", OracleDbType.Varchar2).Value = param.Type;

                    // IsRequired -> -1 / 0
                    cmd.Parameters.Add("IS_REQUIRED", OracleDbType.Int16).Value = param.IsRequired ? -1 : 0;

                    // AllowedValues lo guardamos como JSON si existe
                    if (param.AllowedValues != null && param.AllowedValues.Length > 0)
                    {
                        var jsonAllowed = JsonSerializer.Serialize(param.AllowedValues);
                        cmd.Parameters.Add("ALLOWED_VALUES_JSON", OracleDbType.Clob).Value = jsonAllowed;
                    }
                    else
                    {
                        cmd.Parameters.Add("ALLOWED_VALUES_JSON", OracleDbType.Clob).Value = DBNull.Value;
                    }

                    // BusquedaLike -> -1 / 0 / NULL
                    if (param.BusquedaLike.HasValue)
                    {
                        cmd.Parameters.Add("BUSQUEDA_LIKE", OracleDbType.Int16).Value =
                            param.BusquedaLike.Value ? -1 : 0;
                    }
                    else
                    {
                        cmd.Parameters.Add("BUSQUEDA_LIKE", OracleDbType.Int16).Value = DBNull.Value;
                    }

                    var outParam = new OracleParameter("OUT_PARAM_ID", OracleDbType.Int32)
                    {
                        Direction = ParameterDirection.Output
                    };
                    cmd.Parameters.Add(outParam);

                    await cmd.ExecuteNonQueryAsync();

                    var paramId = Convert.ToInt32(outParam.Value.ToString());

                    // Insertar Values (IntCodeItem) si los hay
                    if (param.Values != null && param.Values.Count > 0)
                    {
                        foreach (var v in param.Values)
                        {
                            await InsertParameterValueAsync(connection, transaction, paramId, v);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante InsertParametersAsync:");
                Console.WriteLine(ex);

            }
        }

        private static async Task InsertParameterValueAsync(
            OracleConnection connection,
            OracleTransaction transaction,
            int paramId,
            IntCodeItem value)
        {
            try
            {
                const string sql = @"
                    INSERT INTO RPT_REPORT_PARAM_VALUE
                        (PARAM_VALUE_ID, PARAM_ID, KEY_INT, VALUE_TEXT)
                    VALUES
                        (RPT_PARAM_VALUE_SEQ.NEXTVAL, :PARAM_ID, :KEY_INT, :VALUE_TEXT)";

                using var cmd = new OracleCommand(sql)
                {
                    Connection = connection,
                    Transaction = transaction
                };

                cmd.Parameters.Add("PARAM_ID", OracleDbType.Int32).Value = paramId;
                cmd.Parameters.Add("KEY_INT", OracleDbType.Int32).Value = value.Key;
                cmd.Parameters.Add("VALUE_TEXT", OracleDbType.Varchar2).Value =
                    (object?)value.Value ?? DBNull.Value;

                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR durante InsertParameterValueAsync:");
                Console.WriteLine(ex);

            }
        }
    }
}

