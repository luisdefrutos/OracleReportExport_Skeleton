using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Configuration;

namespace OracleReportExport.Infrastructure.Services
{
    public sealed class ConnectionCatalogService : IConnectionCatalogService
    {
        private readonly List<ConnectionInfo> _connections;

        public ConnectionCatalogService()
        {
            // <carpeta exe>\Configuration\Connections.json
            var basePath = AppContext.BaseDirectory;
            var configPath = Path.Combine(basePath, "Configuration", "Connections.json");

            if (!File.Exists(configPath))
            {
                throw new FileNotFoundException(
                    $"No se ha encontrado el fichero de conexiones en: {configPath}");
            }

            var json = File.ReadAllText(configPath);

            var root = JsonSerializer.Deserialize<ConnectionConfigRoot>(json)
                       ?? new ConnectionConfigRoot();

            _connections = root.Connections
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

