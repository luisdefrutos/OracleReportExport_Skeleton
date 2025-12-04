using OracleReportExport.Domain.Enums;

namespace OracleReportExport.Domain.Models;

public sealed class ReportDefinition
{
    public string ?Id { get; set; } 
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public ReportSourceType SourceType { get; init; }

    public string? SqlForStations { get; set; }
    public string? SqlForCentral { get; set; }

    public string? SqlFileForStations { get; init; }
    public string? SqlFileForCentral { get; init; }

    public IReadOnlyList<TableMasterParameterDefinition>? TableMasterForParameters { get; set; }
        = Array.Empty<TableMasterParameterDefinition>();

    public IReadOnlyList<ReportParameterDefinition> Parameters { get; init; }
        = Array.Empty<ReportParameterDefinition>();
}
