using System.Collections.Generic;

namespace OracleReportExport.Infrastructure.Configuration
{
    // Representa una conexión individual del JSON
    public sealed class ConnectionConfig
    {
        public string Id { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string ConnectionString { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty; // "Central" o "Estacion"
    }

    // Representa la raíz del JSON: { "Connections": [ ... ] }
    public sealed class ConnectionConfigRoot
    {
        public List<ConnectionConfig> Connections { get; set; } = new();
    }
}
