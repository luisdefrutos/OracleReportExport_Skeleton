using System.Data;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IQueryExecutor
{
    Task<DataTable> ExecuteQueryAsync(
        string sql,
        IReadOnlyDictionary<string, object?> parameters,
        string connectionId,
        string reportId,
        CancellationToken ct = default);
}
