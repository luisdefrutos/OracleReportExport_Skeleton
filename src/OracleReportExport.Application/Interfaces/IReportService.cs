using System.Data;
using OracleReportExport.Domain.Models;

namespace OracleReportExport.Application.Interfaces;

public interface IReportService
{
    Task<IReadOnlyList<ReportDefinition>> GetAvailableReportsAsync(CancellationToken ct = default);

    Task<DataTable> ExecuteReportAsync(
        string reportId,
        IReadOnlyDictionary<string, object?> parameterValues,
        IReadOnlyList<string> targetConnectionIds,
        CancellationToken ct = default);
}
