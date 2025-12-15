using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OracleReportExport.Domain.Enums;

namespace OracleReportExport.Domain.Class
{
    public static class ClassSql
    {
        // Detecta consultas de catálogo Oracle
        // ALL_*, USER_*, DBA_*, V$, GV$
        private static readonly Regex RxCatalogQuery = new Regex(
            @"(?is)\bFROM\s+(ALL_|USER_|DBA_|V\$|GV\$)",
            RegexOptions.Compiled);

        public static SqlKind ClassifySql(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                return SqlKind.Unknown;

            // Normalizar por si viene con BOM
            sql = sql.TrimStart('\uFEFF');

            // 🔹 BLOQUE 1: CATÁLOGO → NO EVALUAR / NO EXPLAIN
            if (RxCatalogQuery.IsMatch(sql))
                return SqlKind.CatalogQuery;

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

            // 🔹 BLOQUE 2: CLASIFICACIÓN NORMAL
            switch (firstToken)
            {
                case "SELECT":
                case "WITH": // WITH también es SQL válido para Explain
                    return SqlKind.Select;

                // DML
                case "UPDATE":
                case "INSERT":
                case "DELETE":
                case "MERGE":
                    return SqlKind.Dml;

                // DDL “poco peligrosa”
                case "ALTER":
                case "CREATE":
                case "RENAME":
                case "COMMENT":
                case "GRANT":
                case "REVOKE":
                    return SqlKind.DdlSafe;

                // DDL peligrosa
                case "DROP":
                case "TRUNCATE":
                    return SqlKind.DdlDangerous;

                // PL/SQL
                case "BEGIN":
                case "DECLARE":
                    return SqlKind.PlSqlBlock;

                default:
                    return SqlKind.Unknown;
            }
        }
    }
}

