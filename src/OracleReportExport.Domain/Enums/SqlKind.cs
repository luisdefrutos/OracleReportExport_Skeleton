using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleReportExport.Domain.Enums
{
    public enum SqlKind
    {
        Unknown,
        Select,
        Dml,           // INSERT / UPDATE / DELETE / MERGE
        DdlSafe,       // ALTER / CREATE / RENAME / COMMENT / GRANT / REVOKE
        DdlDangerous,   // DROP / TRUNCATE
        PlSqlBlock
    }
}
