namespace OracleReportExport.Application.Models;

public sealed class ConnectionInfo
{
    public string Id { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty; // "Central" o "Estacion"
    public override string ToString() =>
          Type == "Central" ? DisplayName : $"{DisplayName} ({Id})";
}
