using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Class;
using System.Data;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IQueryExecutor
{
    Task<DataTable> ExecuteQueryAsync(
        string sql,
        IReadOnlyDictionary<string, object?> parameters,
        ConnectionInfo connectionInfo,
        string reportId,
        CancellationToken ct = default);


    Task<int> ExecuteNonQueryAsync(
      string sql,
      IReadOnlyDictionary<string, object?> parameters,
      ConnectionInfo connectionInfo,
      string reportId,
      CancellationToken ct = default);

    Task<bool> ValidateSqlSyntaxAsync(
      string sql,
      ConnectionInfo connection,
      CancellationToken ct);
}
