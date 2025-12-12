using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OracleReportExport.Domain.Enums;

namespace OracleReportExport.Domain.Class
{
    public static class ClassSql
    {
        public static SqlKind ClassifySql(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                return SqlKind.Unknown;

            var lines = sql
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim());

            string? firstMeaningfulLine = lines
                .FirstOrDefault(l => !l.StartsWith("--") && !l.StartsWith("/*"));

            if (string.IsNullOrWhiteSpace(firstMeaningfulLine))
                firstMeaningfulLine = sql.Trim();

            var parts = firstMeaningfulLine
                .TrimStart()
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
                return SqlKind.Unknown;

            var firstToken = parts[0].ToUpperInvariant();

         

            switch (firstToken)
            {
                case "SELECT":
                    return SqlKind.Select;

                // DML
                case "UPDATE":
                case "INSERT":
                case "DELETE":
                case "MERGE":
                    return SqlKind.Dml;

                // DDL “poco peligrosa”: ALTER, CREATE, etc.
                case "ALTER":
                case "CREATE":
                case "RENAME":
                case "COMMENT":
                case "GRANT":
                case "REVOKE":
                    return SqlKind.DdlSafe;

                // DDL peligrosa: DROP / TRUNCATE
                case "DROP":
                case "TRUNCATE":
                    return SqlKind.DdlDangerous;

                case "BEGIN":
                case "DECLARE":
                    return SqlKind.PlSqlBlock;

                default:
                    return SqlKind.Unknown;
            }
        }

    }
}
