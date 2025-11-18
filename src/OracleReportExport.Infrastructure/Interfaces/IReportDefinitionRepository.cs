using OracleReportExport.Domain.Models;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IReportDefinitionRepository
{
    Task<IReadOnlyList<ReportDefinition>> GetAllAsync(CancellationToken ct = default);
    Task<ReportDefinition?> GetByIdAsync(string id, CancellationToken ct = default);
}
