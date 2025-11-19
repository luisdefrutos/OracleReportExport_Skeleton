using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text.Json;
using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Infrastructure.Configuration;
using OracleReportExport.Infrastructure.Interfaces;

namespace OracleReportExport.Infrastructure.Data
{
    public sealed class OracleConnectionFactory : IOracleConnectionFactory
    {
        private readonly Dictionary<string, ConnectionConfig> _connections;

        public OracleConnectionFactory()
        {
            // Ruta del JSON en la carpeta del ejecutable:
            // <carpeta exe>\Configuration\Connections.json
            var basePath = AppContext.BaseDirectory;
            var configPath = Path.Combine(basePath, "Configuration", "Connections.json");



            if (!File.Exists(configPath))
                throw new FileNotFoundException($"No se encontró el fichero de conexiones: {configPath}");

            var json = File.ReadAllText(configPath);
            
            

            var root = JsonSerializer.Deserialize<ConnectionConfigRoot>(json)
                       ?? new ConnectionConfigRoot();

            // Diccionario: Id  -> conexión
            _connections = root.Connections
                               .ToDictionary(c => c.Id,
                                             c => c,
                                             StringComparer.OrdinalIgnoreCase);

            if(File.Exists(configPath))
                File.Delete(configPath);
        }

        public DbConnection CreateConnection(string connectionId)
        {
            if (!_connections.TryGetValue(connectionId, out var cfg))
                throw new ArgumentException(
                    $"No existe la conexión '{connectionId}' en Connections.json",
                    nameof(connectionId));

            // Creamos la conexión Oracle pero NO la abrimos.
            var conectionçactive= new OracleConnection(cfg.ConnectionString);
            return conectionçactive;
        }
    }
}

