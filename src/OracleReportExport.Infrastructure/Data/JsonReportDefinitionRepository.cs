using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Configuration;
using OracleReportExport.Infrastructure.Interfaces;

namespace OracleReportExport.Infrastructure.Data
{
    public sealed class JsonReportDefinitionRepository : IReportDefinitionRepository
    {
        private readonly List<ReportDefinition> _reports;

        public JsonReportDefinitionRepository()
        {
            var basePath = AppContext.BaseDirectory;
            var path = Path.Combine(basePath, "Configuration", "Reports.json");
            if (!File.Exists(path))
                throw new FileNotFoundException($"No se encontró Reports.json en: {path}");

            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            options.Converters.Add(new JsonStringEnumConverter());

            var root = JsonSerializer.Deserialize<ReportDefinitionRoot>(json, options)
                       ?? new ReportDefinitionRoot();
            var processedReports = new List<ReportDefinition>();
            foreach (var r in root.Reports)
            {
                var report = r;
                if (!string.IsNullOrWhiteSpace(r.SqlFileForStations))
                {
                    var sqlPath = Path.Combine(basePath, r.SqlFileForStations);
                    if (!File.Exists(sqlPath))
                        throw new FileNotFoundException($"No se encontró el archivo SQL: {sqlPath}");
                    var sqlText = File.ReadAllText(sqlPath);
                    // Creamos una nueva instancia con la SQL cargada
                    report = new ReportDefinition
                    {
                        Id = r.Id,
                        Name = r.Name,
                        Category = r.Category,
                        Description = r.Description,
                        SourceType = r.SourceType,
                        SqlForStations = sqlText,
                        SqlForCentral = r.SqlForCentral,
                        SqlFileForStations = r.SqlFileForStations,
                        SqlFileForCentral = r.SqlFileForCentral,
                        Parameters = r.Parameters,
                       TableMasterForParameters=r.TableMasterForParameters
                    };
                }
                else if (!string.IsNullOrWhiteSpace(r.SqlFileForCentral))
                {
                    var sqlPath = Path.Combine(basePath, r.SqlFileForCentral);
                    if (!File.Exists(sqlPath))
                        throw new FileNotFoundException($"No se encontró el archivo SQL: {sqlPath}");
                    var sqlText = File.ReadAllText(sqlPath);
                    report = new ReportDefinition
                    {
                        Id = r.Id,
                        Name = r.Name,
                        Category = r.Category,
                        Description = r.Description,
                        SourceType = r.SourceType,
                        SqlForStations = r.SqlForStations,
                        SqlForCentral = sqlText,
                        SqlFileForStations = r.SqlFileForStations,
                        SqlFileForCentral = r.SqlFileForCentral,
                        Parameters = r.Parameters,
                        TableMasterForParameters = r.TableMasterForParameters
                    };
                }
                processedReports.Add(report);
            }
            _reports = processedReports;
        }

        public Task<IReadOnlyList<ReportDefinition>> GetAllAsync(CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<ReportDefinition>>(_reports);

        public Task<ReportDefinition?> GetByIdAsync(string id, CancellationToken ct = default)
        {
            var report = _reports
                .FirstOrDefault(r => string.Equals(r.Id, id, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult(report);
        }
    }
}
