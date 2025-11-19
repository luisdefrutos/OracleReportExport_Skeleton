namespace OracleReportExport.Domain.Models;

public sealed class ReportParameterDefinition
{
    public string Name { get; init; } = string.Empty;
    public string Label { get; init; } = string.Empty;
    public string Type { get; init; } = "string"; // "string", "int", "decimal", "date"
    public bool IsRequired { get; init; }
    public string[]? AllowedValues { get; init; }
    public List<IntCodeItem> Values { get; set; } = new();
}
