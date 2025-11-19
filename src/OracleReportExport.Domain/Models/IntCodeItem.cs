using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleReportExport.Domain.Models
{
    public class IntCodeItem
    {
        public int Key { get; set; }      // 0 ó 1
        public string? Value { get; set; } // "S" ó "N"
    }

}
