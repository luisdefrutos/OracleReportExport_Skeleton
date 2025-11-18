using OracleReportExport.Application.Models;

namespace OracleReportExport.Application.Interfaces;

public interface IConnectionCatalogService
{
    IReadOnlyList<ConnectionInfo> GetAllConnections();
}
