using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleReportExport.Domain.Models
{
    public sealed class ReportQueryResult
    {
        public DataTable? Data { get; init; }
        public IReadOnlyList<string> TimeoutConnections { get; init; } = Array.Empty<string>();
    }
}
