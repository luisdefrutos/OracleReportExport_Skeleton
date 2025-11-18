using System.Data;

namespace OracleReportExport.Application.Interfaces;

public sealed class AdHocValidationResult
{
    public bool IsValid { get; init; }
    public bool ContainsDangerousCommands { get; init; }
    public bool ContainsDml { get; init; }
    public string Message { get; init; } = string.Empty;
}

public interface IAdHocQueryService
{
    AdHocValidationResult ValidateSql(string sql);

    Task<DataTable> ExecuteSqlAsync(
        string sql,
        string connectionId,
        CancellationToken ct = default);
}
