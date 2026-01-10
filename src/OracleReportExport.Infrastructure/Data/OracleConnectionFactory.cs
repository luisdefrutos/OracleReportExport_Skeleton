using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Configuration;
using OracleReportExport.Infrastructure.Interfaces;
using OracleReportExport.Infrastructure.Services;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace OracleReportExport.Infrastructure.Data
{
    public sealed class OracleConnectionFactory : IOracleConnectionFactory
    {
        private readonly Dictionary<string, ConnectionConfig> _connections;
        private readonly ConnectionCatalogService _connectionCatalog;
        private readonly object _syncLock = new();

        // Propiedades/operaciones para añadir entradas de forma segura.
        // Uso preferible: métodos públicos en lugar de propiedades write-only.
        public void AddOrUpdateConnection(string key, ConnectionConfig cfg)
        {
            if (string.IsNullOrWhiteSpace(key) || cfg == null)
                return;

            lock (_syncLock)
            {
                _connections[key] = cfg;
            }
        }

        public void AddOrUpdateConnection(ConnectionInfo info)
        {
            if (info == null)
                return;

            var key = string.Concat(info.Id, "_", info.DisplayName);
            var cfg = new ConnectionConfig
            {
                Id = info.Id,
                DisplayName = info.DisplayName,
                ConnectionString = info.ConnectionString,
                Type = info.Type
            };

            AddOrUpdateConnection(key, cfg);
        }

        public void AddOrUpdateConnections(IEnumerable<KeyValuePair<string, ConnectionConfig>> items)
        {
            if (items == null)
                return;

            lock (_syncLock)
            {
                foreach (var kv in items)
                {
                    if (string.IsNullOrWhiteSpace(kv.Key) || kv.Value == null)
                        continue;

                    _connections[kv.Key] = kv.Value;
                }
            }
        }

        public void AddOrUpdateConnections(IEnumerable<ConnectionInfo> infos)
        {
            if (infos == null)
                return;

            lock (_syncLock)
            {
                foreach (var info in infos)
                {
                    if (info == null)
                        continue;

                    var key = string.Concat(info.Id, "_", info.DisplayName);
                    _connections[key] = new ConnectionConfig
                    {
                        Id = info.Id,
                        DisplayName = info.DisplayName,
                        ConnectionString = info.ConnectionString,
                        Type = info.Type
                    };
                }
            }
        }

        public OracleConnectionFactory(ConnectionCatalogService connectionCatalog)
        {
            _connectionCatalog = connectionCatalog;
            _connections = new Dictionary<string, ConnectionConfig>();

            foreach (ConnectionInfo itemConnection in _connectionCatalog.GetAllConnections())
            {
                _connections.Add(String.Concat(itemConnection.Id, "_", itemConnection.DisplayName), new ConnectionConfig()
                {
                    ConnectionString = itemConnection.ConnectionString,
                    DisplayName = itemConnection.DisplayName,
                    Id = itemConnection.Id,
                    Type = itemConnection.Type,
                });
            }
           
        }

        public DbConnection CreateConnection(string connectionId)
        {
            if (!_connections.TryGetValue(connectionId, out var cfg))
                throw new ArgumentException(
                    $"No existe la conexión '{connectionId}' ",
                    nameof(connectionId));

            // Creamos la conexión Oracle pero NO la abrimos.
            var conectionactive = new OracleConnection(cfg.ConnectionString);
            return conectionactive;
        }
    }
}

