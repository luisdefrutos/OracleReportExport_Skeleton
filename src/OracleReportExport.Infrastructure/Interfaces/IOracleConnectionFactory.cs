using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Configuration;
using System.Collections.Generic;
using System.Data.Common;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IOracleConnectionFactory
{
    DbConnection CreateConnection(string connectionId);

    // Añadir/actualizar una entrada por clave
    void AddOrUpdateConnection(string key, ConnectionConfig cfg);

    // Añadir/actualizar a partir de ConnectionInfo
    void AddOrUpdateConnection(ConnectionInfo info);

    // Añadir/actualizar varias entradas a partir de pares key/Configuration
    void AddOrUpdateConnections(IEnumerable<KeyValuePair<string, ConnectionConfig>> items);

    // Añadir/actualizar varias entradas a partir de ConnectionInfo
    void AddOrUpdateConnections(IEnumerable<ConnectionInfo> infos);
}
