using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace OracleReportExport.Infrastructure.Services
{
    public sealed class ConnectionCatalogService : IConnectionCatalogService
    {
        private readonly List<ConnectionInfo> _connections;

        public ConnectionCatalogService()
        {
            var basePath = AppContext.BaseDirectory;

            // Read central connection definition
            var centralConfigPath = Path.Combine(basePath, "Configuration", "ConnectionsCentral.json");
            if (!File.Exists(centralConfigPath))
                throw new FileNotFoundException($"No se ha encontrado el fichero de conexiones central en: {centralConfigPath}");

            var centralJson = File.ReadAllText(centralConfigPath);
            var centralRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(centralJson)
                              ?? new ConnectionConfigRoot();

            // Ensure we have at least the central entry
            var centralEntry = centralRoot.Connections.FirstOrDefault();
            if (centralEntry is null)
                throw new InvalidOperationException("ConnectionsCentral.json no contiene ninguna entrada de conexión central.");

            // List that will contain central + estaciones
            var connectionsList = new List<ConnectionConfig>();

            // Ensure central entry has Type = "Central"
            centralEntry.Type = "Central";
            connectionsList.Add(centralEntry);

            // Now try to query the central DB to obtain estaciones
            try
            {
                using var conn = new OracleConnection(centralEntry.ConnectionString);
                conn.Open();

                const string sql = @"SELECT CODESTACION, DESESTACION, CONNECTION_STRING
                                    FROM INSTANCIAS_BBDD
                                    WHERE
                                    ACTIVO = -1
                                    AND DESESTACION LIKE '%I.T.V.%'
                                    ORDER BY 1";

                using var cmd = conn.CreateCommand();
                cmd.CommandText = sql;

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var id = reader.IsDBNull(0) ? string.Empty : reader.GetString(0).Trim();
                    var displayName = reader.IsDBNull(1) ? ("Estación " + id) : reader.GetString(1).Trim();
                    var connStr = reader.IsDBNull(2) ? string.Empty : reader.GetString(2).Trim();

                    // Apply replacements requested by user
                    connStr = connStr
                        .Replace("UID=", "User Id=")
                        .Replace("PWD=", "Password=")
                        .Replace("SERVER=", "Data Source=");

                    // If the returned connection string looks like it lacks the User Id/Password tokens and
                    // the project examples use User Id=ITV;Password=ItvitV.-20; we will NOT inject defaults here
                    // to avoid accidental credentials. The existing connection_string should be complete.

                    if (string.IsNullOrWhiteSpace(id))
                        continue;

                    connectionsList.Add(new ConnectionConfig
                    {
                        Id = id,
                        DisplayName = displayName,
                        ConnectionString = connStr,
                        Type = "Estacion"
                    });
                }
            }
            catch (Exception)
            {
                // If querying central fails, fall back to existing Connections.json if present.
                // We'll try to read Configuration\Connections.json and use that. If that also fails,
                // we'll keep only the central entry constructed above.
                var fallbackPath = Path.Combine(basePath, "Configuration", "Connections.json");
                if (File.Exists(fallbackPath))
                {
                    try
                    {
                        var fallbackJson = File.ReadAllText(fallbackPath);
                        var fallbackRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(fallbackJson)
                                           ?? new ConnectionConfigRoot();

                        // Merge: ensure central is first, then unique demás
                        foreach (var c in fallbackRoot.Connections)
                        {
                            // skip central duplicate (we already have it)
                            if (string.Equals(c.Id, centralEntry.Id, StringComparison.OrdinalIgnoreCase))
                                continue;

                            connectionsList.Add(c);
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            // Write the updated Connections.json so the rest of the application can use it
            try
            {
                var outputRoot = new ConnectionConfigRoot { Connections = connectionsList };
                var options = new JsonSerializerOptions { WriteIndented = true };

                var outputPath = Path.Combine(basePath, "Configuration", "Connections.json");
                var outJson = JsonSerializer.Serialize(outputRoot, options);
                File.WriteAllText(outputPath, outJson);
            }
            catch
            {
                // best-effort write; if it fails, continue with in-memory list
            }

            // Populate the _connections list exposed by the service
            _connections = connectionsList
                .Select(c => new ConnectionInfo
                {
                    Id = c.Id,
                    DisplayName = c.DisplayName,
                    Type = c.Type
                })
                .OrderBy(c => c.Type)
                .ThenBy(c => c.Id)
                .ToList();
        }

        public IReadOnlyList<ConnectionInfo> GetAllConnections()
            => _connections;
    }
}

