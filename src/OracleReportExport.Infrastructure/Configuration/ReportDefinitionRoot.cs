using System.Collections.Generic;
using OracleReportExport.Domain.Models;

namespace OracleReportExport.Infrastructure.Configuration
{
    public class ReportDefinitionRoot
    {
        public List<ReportDefinition> Reports { get; set; } = new();
    }
}
