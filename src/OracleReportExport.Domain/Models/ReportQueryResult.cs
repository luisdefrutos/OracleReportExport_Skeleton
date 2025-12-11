using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OracleReportExport.Domain.Enums;


namespace OracleReportExport.Domain.Models
{
    public sealed class ReportQueryResult
    {
         public SqlKind Kind { get; init; }
        public DataTable? Data { get; init; }
        public IReadOnlyList<string> TimeoutConnections { get; init; } = Array.Empty<string>();
        public int RowsAffected { get; init; }
    }
}
