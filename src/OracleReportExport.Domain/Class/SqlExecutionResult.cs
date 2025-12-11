using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OracleReportExport.Domain.Enums;

namespace OracleReportExport.Domain.Class
{
    public sealed class SqlExecutionResult
    {
        public SqlKind Kind { get; init; }          // Select, Dml, Ddl, Unknown
        public DataTable? Data { get; init; }       // Solo para SELECT
        public int RowsAffected { get; init; }      // Para DML/DDL
    }

}
