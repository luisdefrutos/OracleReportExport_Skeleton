using System.Data.Common;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IOracleConnectionFactory
{
    DbConnection CreateConnection(string connectionId);
}
