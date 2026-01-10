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
    public sealed class ConnectionCatalogService : IConnectionCatalogService
    {
        public readonly List<ConnectionInfo> _connections;


        public ConnectionInfo? AddConnection
        {
            set
            {
                if (value == null)
                    return;

                _connections?.Add(value);
            }
        }
        public IEnumerable<ConnectionInfo>? AddConnections
        {
            set
            {
                if (value == null || _connections == null)
                    return;

                foreach (var connection in value)
                {
                    if (connection != null)
                        _connections.Add(connection);
                }
            }
        }

        public ConnectionCatalogService()
        {
            var basePath = AppContext.BaseDirectory;
            var centralConfigPath = Path.Combine(basePath, "Configuration", "ConnectionsCentral.json");
            if (!File.Exists(centralConfigPath))
                throw new FileNotFoundException($"No se ha encontrado el fichero de conexiones central en: {centralConfigPath}");
            var centralJson = File.ReadAllText(centralConfigPath);
            var centralRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(centralJson)?? new ConnectionConfigRoot();
            var connectionsList = new List<ConnectionConfig>();
            try
            {
                var initialJson = File.ReadAllText(centralConfigPath);
                var initialStationRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(initialJson) ?? new ConnectionConfigRoot();
                foreach (var c in initialStationRoot.Connections)
                {
                    c.ConnectionString = c.ConnectionString
                             .Replace("UID=", "User Id=")
                             .Replace("PWD=", "Password=")
                            .Replace("SERVER=", "Data Source=");
                    connectionsList.Add(new ConnectionConfig
                    {
                        Id = c.Id,
                        DisplayName = c.DisplayName,
                        ConnectionString = c.ConnectionString,
                        Type = c.Type
                    });
                }
                _connections = connectionsList
                  .Select(c => new ConnectionInfo
                  {
                      Id = c.Id,
                      DisplayName = c.DisplayName,
                      Type = c.Type,
                      ConnectionString = c.ConnectionString
                  })
                  .OrderBy(c => c.Type)
                  .ThenBy(c => c.Id)
                  .ToList();
            }
            catch (Exception ex)
            {
                var fallbackPath = Path.Combine(basePath, "Configuration", "ConnectionsCentral.json");
                if (File.Exists(fallbackPath))
                {
                    try
                    {
                        var fallbackJson = File.ReadAllText(fallbackPath);
                        var fallbackRoot = JsonSerializer.Deserialize<ConnectionConfigRoot>(fallbackJson)
                                           ?? new ConnectionConfigRoot();
                        foreach (var c in fallbackRoot.Connections)
                        {
                            connectionsList.Add(c);
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
        }

        public IReadOnlyList<ConnectionInfo> GetAllConnections()
            => _connections;

    }
}

