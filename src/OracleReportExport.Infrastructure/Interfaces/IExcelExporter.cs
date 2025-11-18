using System.Data;

namespace OracleReportExport.Infrastructure.Interfaces;

public interface IExcelExporter
{
    void ExportToExcel(DataTable table, string filePath);
}
