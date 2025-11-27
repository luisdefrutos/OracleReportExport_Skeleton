using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OracleReportExport.Presentation.Desktop
{
    public static class GlobalExceptionHandler
    {
        public static void Initialize()
        {
            // Excepciones de hilos de UI (WinForms)
            System.Windows.Forms.Application.ThreadException += Application_ThreadException;

            // Excepciones NO controladas en otros hilos / tareas
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            // Para tareas async "fire and forget"
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
            => Handle(e.Exception, "UI Thread");

        private static void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
            => Handle(e.ExceptionObject as Exception, "AppDomain");

        private static void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            Handle(e.Exception, "TaskScheduler");
            e.SetObserved(); // Evita que termine el proceso
        }

        /// <summary>
        /// Centro único para tratar / loguear las excepciones.
        /// </summary>
        public static void Handle(Exception? ex, string? context = null, string? messageAdHoc = null)
        {
            if (ex == null) return;

            try
            {
                // Aquí podrías meter Serilog, NLog, escribir a fichero, etc.
                // Log.Fatal(ex, "Error global en {Context}", context);
                var messageSettings = String.IsNullOrWhiteSpace(messageAdHoc) ? ex.Message : messageAdHoc;
                var msgContext = string.IsNullOrWhiteSpace(context) ? "" : $"[{context}] ";
                var message = $"{msgContext}Se ha producido un error.\n" +
                              $"{messageSettings}\n" +
                              "Si el problema persiste, contacte con el soporte.";

                MessageBox.Show(message,
                    "Error inesperado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch
            {
                
            }
        }
    }

}
