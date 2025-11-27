using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Interfaces;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
 

namespace OracleReportExport.Infrastructure.Data
{
    public sealed class OracleQueryExecutor : IQueryExecutor
    {
        private readonly IOracleConnectionFactory _connectionFactory;

        public OracleQueryExecutor(IOracleConnectionFactory connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        public async Task<DataTable> ExecuteQueryAsync(
            string sql,
            IReadOnlyDictionary<string, object?> parameters,
             ConnectionInfo  connectionInfo,
            string reportId,
            CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(sql))
                throw new ArgumentException("La SQL no puede estar vacía.", nameof(sql));

            var result = new DataTable();

            await using var conn = _connectionFactory.CreateConnection(String.Concat(connectionInfo.Id,"_",connectionInfo.DisplayName)) as OracleConnection
                                   ?? throw new InvalidOperationException("La conexión devuelta no es OracleConnection.");

            await conn.OpenAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.BindByName = true;
            cmd.CommandText = sql;

            // Parámetros opcionales
            if (parameters is not null)
            {
                foreach (var kvp in parameters)
                {
                    var param = cmd.CreateParameter();
                    // En Oracle suelen ir con :, pero por si acaso lo añadimos si falta
                    param.ParameterName = kvp.Key.StartsWith(":", StringComparison.Ordinal)
                        ? kvp.Key
                        : ":" + kvp.Key;

                    param.Value = kvp.Value ?? DBNull.Value;
                    cmd.Parameters.Add(param);
                }
            }


            var debugSql = BuildDebugSql(cmd);

            Log.Information($"Ejecutando SQL de '{reportId}' en {connectionInfo.ToString()}:\n{debugSql}");

            using var registration = ct.Register(() => cmd.Cancel());

            using var reader = await cmd.ExecuteReaderAsync(ct);
            result.Load(reader);

            return result;
        }
        private static string BuildDebugSql(OracleCommand cmd)
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine(cmd.CommandText);
            string txtReplaced=sb.ToString();
            foreach (OracleParameter p in cmd.Parameters)
            {
                txtReplaced= txtReplaced.ToString().Replace(p.ParameterName, FormatParameterValue(p.Value));
            }
            sb.Clear();
            sb.AppendLine(txtReplaced);
            return sb.ToString();
            }

    private static string FormatParameterValue(object? value)
    {
        if (value is null || value == DBNull.Value)
            return "NULL";

        switch (value)
        {
            case DateTime dt:
                // Representación Oracle amigable
                return $"TO_DATE('{dt:dd/MM/yyyy HH:mm:ss}', 'dd/mm/yyyy hh24:mi:ss')";

            case string s:
                // Escapar comillas simples
                return $"'{s.Replace("'", "''")}'";

            case bool b:
                return b ? "1" : "0";

            case int or long or short or decimal or double or float:
                // Usar punto como separador decimal
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "NULL";

            default:
                return $"'{value.ToString()?.Replace("'", "''")}'";
        }
    }

        public async Task<bool> ValidateSqlSyntaxAsync(string sql,
                 ConnectionInfo connection,
                 CancellationToken ct)
                    {
                        using var conn = (OracleConnection)_connectionFactory.CreateConnection(
                            string.Concat(connection.Id, "_", connection.DisplayName));

                        await conn.OpenAsync(ct);

                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = $"EXPLAIN PLAN FOR {sql}";

                        using var registration = ct.Register(() => cmd.Cancel());

                       var result= await cmd.ExecuteNonQueryAsync(ct);
                          return result >= 0;
            // Si la sintaxis es mala -> OracleException ORA-009xx
        }

    }
}
