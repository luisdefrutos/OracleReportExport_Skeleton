using Serilog;
using System;
using System.Windows.Forms;

namespace OracleReportExport.Presentation.Desktop
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {

            var logPath = Path.Combine(AppContext.BaseDirectory, "logs");
            Directory.CreateDirectory(logPath);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    Path.Combine(logPath, "app-.log"),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 30,
                    encoding: System.Text.Encoding.UTF8)
                .CreateLogger();
            // Estilo WinForms clásico
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            GlobalExceptionHandler.Initialize();

            System.Windows.Forms.Application.Run(new MainForm());
        }
    }
}


