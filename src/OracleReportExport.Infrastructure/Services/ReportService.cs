using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
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
                    if (!string.IsNullOrWhiteSpace(report.SqlFileForCentral))
                    {
                        // asume ruta relativa dentro de tu Configuration/SQL...
                        var basePath = AppContext.BaseDirectory;
                        var pathSql = Path.Combine(basePath, report.SqlFileForCentral);
                        if (!File.Exists(pathSql))
                            throw new FileNotFoundException($"No se encontró el fichero SQL: {pathSql}");
                        sql = await File.ReadAllTextAsync(pathSql);
                    }
                    else
                    {
                        sql = report.SqlForCentral ?? string.Empty;
                    }
                    break;
                case OracleReportExport.Domain.Enums.ReportSourceType.Estacion:
                    if (!string.IsNullOrWhiteSpace(report.SqlFileForStations))
                    {
                        var basePath = AppContext.BaseDirectory;
                        var pathSql = Path.Combine(basePath, report.SqlFileForStations);

                        if (!File.Exists(pathSql))
                            throw new FileNotFoundException($"No se encontró el fichero SQL: {pathSql}");

                        sql = await File.ReadAllTextAsync(pathSql);
                    }
                    else
                    {
                        sql = report.SqlForStations ?? string.Empty;
                    }
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

   }
}

