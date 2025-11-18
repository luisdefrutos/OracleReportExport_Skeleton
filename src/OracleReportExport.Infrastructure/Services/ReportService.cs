using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Interfaces;

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

        public async Task<DataTable> ExecuteReportAsync(
            string reportId,
            IReadOnlyDictionary<string, object?> parameterValues,
            IReadOnlyList<string> targetConnectionIds,
            CancellationToken ct = default)
        {
            var def = await _definitions.GetByIdAsync(reportId, ct);

            if (def is null)
                throw new InvalidOperationException(
                    $"No se encontró la definición del informe '{reportId}'.");

            var sql = def.SqlForStations ?? def.SqlForCentral;

            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException(
                    $"El informe '{reportId}' no tiene SQL configurada.");

            DataTable? combined = null;

            foreach (var connectionId in targetConnectionIds)
            {
                var table = await _queryExecutor.ExecuteQueryAsync(
                    sql,
                    parameterValues,
                    connectionId,
                    reportId,
                    ct);

                if (combined is null)
                {
                    combined = table.Clone();
                    combined.Columns.Add("CONEXION_ID", typeof(string));
                }

                foreach (DataRow row in table.Rows)
                {
                    var newRow = combined.NewRow();

                    foreach (DataColumn col in table.Columns)
                    {
                        newRow[col.ColumnName] = row[col];
                    }

                    newRow["CONEXION_ID"] = connectionId;
                    combined.Rows.Add(newRow);
                }
            }

            return combined ?? new DataTable();
        }
    }
}

