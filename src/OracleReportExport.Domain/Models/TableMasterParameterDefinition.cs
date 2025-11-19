using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleReportExport.Domain.Models
{
    public class TableMasterParameterDefinition
    {

        public string? Name { get; set; }
        public string? Label { get; set; }
        public string? Type { get; set; }
        public bool IsRequired { get; set; }
        public string? Id { get; set; }
        public string? Text { get; set; }
        public List<string?>? ValuesRequired { get; set; }

        public string? SqlQueryMaster { get; set; }
    }
 
}
