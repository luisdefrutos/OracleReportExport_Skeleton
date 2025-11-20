using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Models;
using System.Data;

namespace OracleReportExport.Application.Interfaces;

public interface IReportService
{
    Task<IReadOnlyList<ReportDefinition>> GetAvailableReportsAsync(CancellationToken ct = default);

      Task<ReportQueryResult> ExecuteReportAsync(
        ReportDefinition report,
        IReadOnlyDictionary<string, object?> parameterValues,
        List<ConnectionInfo> targetConnection,
        CancellationToken ct = default);
}
