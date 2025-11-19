using OracleReportExport.Domain.Enums;

namespace OracleReportExport.Domain.Models;

public sealed class ReportDefinition
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public ReportSourceType SourceType { get; init; }

    public string? SqlForStations { get; init; }
    public string? SqlForCentral { get; init; }

    public string? SqlFileForStations { get; init; }
    public string? SqlFileForCentral { get; init; }

    public IReadOnlyList<TableMasterParameterDefinition>? TableMasterForParameters { get; set; }
        = Array.Empty<TableMasterParameterDefinition>();

    public IReadOnlyList<ReportParameterDefinition> Parameters { get; init; }
        = Array.Empty<ReportParameterDefinition>();
}
