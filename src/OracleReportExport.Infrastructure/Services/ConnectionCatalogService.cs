using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Configuration;
using OracleReportExport.Infrastructure.Data;
using OracleReportExport.Infrastructure.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace OracleReportExport.Infrastructure.Services
{
    public   sealed   class ConnectionCatalogService : IConnectionCatalogService
    {
        private readonly List<ConnectionInfo> _connections;

         public    ConnectionCatalogService()
        {
            var basePath = AppContext.BaseDirectory;
            var centralConfigPath = Path.Combine(basePath, "Configuration", "ConnectionsCentral.json");
            if (!File.Exists(centralConfigPath))
                throw new FileNotFoundException($"No se ha encontrado el fichero de conexiones central en: {centralConfigPath}");

            var centralJson = File.ReadAllText(centralConfigPath);
            var centralRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(centralJson)
                              ?? new ConnectionConfigRoot();
            var centralEntry = centralRoot.Connections.FirstOrDefault();
            if (centralEntry is null)
                throw new InvalidOperationException("ConnectionsCentral.json no contiene ninguna entrada de conexión central.");
            var connectionsList = new List<ConnectionConfig>();
           centralEntry.Type = "Central";
            connectionsList.Add(centralEntry);

           var initialConfigPath = Path.Combine(basePath, "Configuration", "Connections.json");
            if (!File.Exists(initialConfigPath))
                throw new FileNotFoundException($"No se ha encontrado inicial de conexion en: {initialConfigPath}");

            var initialJson = File.ReadAllText(initialConfigPath);
            var initialStationRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(initialJson)
                              ?? new ConnectionConfigRoot();
          var conectionInitial = initialStationRoot.Connections.Where(x => x.Type == "Central").FirstOrDefault();
           
            try
            {
                 var reportDef = new JsonReportDefinitionRepository();
                var reporItemDefinition = reportDef.GetAllAsync().Result.Where(x => x.Name.Contains("Listado Estaciones Activas")).FirstOrDefault();
                using var conn = new OracleConnection(conectionInitial.ConnectionString);
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = (reporItemDefinition?.SqlForCentral);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var id = reader.IsDBNull(0) ? string.Empty : reader.GetString(0).Trim();
                    var displayName = reader.IsDBNull(1) ? ("Estación " + id) : reader.GetString(1).Trim();
                    var connStr = reader.IsDBNull(2) ? string.Empty : reader.GetString(2).Trim();
                    connStr = connStr
                        .Replace("UID=", "User Id=")
                        .Replace("PWD=", "Password=")
                        .Replace("SERVER=", "Data Source=");
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
            catch (Exception ex)
            {
                var fallbackPath = Path.Combine(basePath, "Configuration", "Connections.json");
                if (File.Exists(fallbackPath))
                {
                    try
                    {
                        var fallbackJson = File.ReadAllText(fallbackPath);
                        var fallbackRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(fallbackJson)
                                           ?? new ConnectionConfigRoot();
                       foreach (var c in fallbackRoot.Connections)
                        {
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

           try
            {
                var outputRoot = new ConnectionConfigRoot { Connections = connectionsList };
                var options = new JsonSerializerOptions { WriteIndented = true };
                var outputPath = Path.Combine(basePath, "Configuration", "Connections.json");
                var outJson = JsonSerializer.Serialize(outputRoot, options);
                File.WriteAllText(outputPath, outJson);
                //if(File.Exists(centralConfigPath))
                //    File.Delete(centralConfigPath);
            }
            catch
            {
            }
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

