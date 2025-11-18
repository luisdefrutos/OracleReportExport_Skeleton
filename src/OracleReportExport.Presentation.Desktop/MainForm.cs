//using OracleReportExport.Application.Models;
//using OracleReportExport.Infrastructure.Data;
//using OracleReportExport.Infrastructure.Interfaces;
//using OracleReportExport.Infrastructure.Services;
//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Linq;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//namespace OracleReportExport.Presentation.Desktop
//{
//    public class MainForm : Form
//    {
//        private readonly TabControl _tabControl = new();

//        private readonly DataGridView _grid = new()
//        {
//            Dock = DockStyle.Fill,
//            ReadOnly = true,
//            AllowUserToAddRows = false
//        };

//        // Lista de conexiones (Central + estaciones)
//        private readonly CheckedListBox _chkConnections = new()
//        {
//            Dock = DockStyle.Left,
//            Width = 260,
//            CheckOnClick = true
//        };

//        // Botones de acciones (con AutoSize para que se vea bien el texto)
//        private readonly Button _btnSelectAll = new()
//        {
//            Text = "Marcar todas",
//            AutoSize = true
//        };

//        private readonly Button _btnUnselectAll = new()
//        {
//            Text = "Desmarcar",
//            AutoSize = true
//        };

//        private readonly Button _btnExport = new()
//        {
//            Text = "Exportar a Excel",
//            AutoSize = true
//        };

//        private readonly Button _btnTest = new()
//        {
//            Text = "Probar consulta",
//            AutoSize = true
//        };

//        // Panel superior para los botones
//        private readonly FlowLayoutPanel _topPanel = new()
//        {
//            Dock = DockStyle.Top,
//            Height = 40,
//            FlowDirection = FlowDirection.LeftToRight,
//            WrapContents = false
//        };

//        // Servicios
//        private readonly ConnectionCatalogService _connectionCatalog;
//        private readonly IQueryExecutor _queryExecutor;

//        public MainForm()
//        {
//            Text = "Oracle Report Export";
//            Width = 1200;
//            Height = 800;

//            _tabControl.Dock = DockStyle.Fill;

//            var tabPredefinidos = new TabPage("Informes predefinidos");
//            var tabAdHoc = new TabPage("SQL avanzada");

//            // Inicializar servicios
//            _connectionCatalog = new ConnectionCatalogService();

//            IOracleConnectionFactory connectionFactory = new OracleConnectionFactory();
//            _queryExecutor = new OracleQueryExecutor(connectionFactory);

//            // Cargar conexiones en el CheckedListBox
//            CargarConexiones();

//            // Configurar barra superior de botones
//            _topPanel.Controls.Add(_btnSelectAll);
//            _topPanel.Controls.Add(_btnUnselectAll);
//            _topPanel.Controls.Add(_btnExport);
//            _topPanel.Controls.Add(_btnTest);

//            // Eventos de botones
//            _btnSelectAll.Click += (_, __) => SetAllConnectionsChecked(true);
//            _btnUnselectAll.Click += (_, __) => SetAllConnectionsChecked(false);
//            _btnTest.Click += async (_, __) => await ProbarConsultaAsync();
//            // _btnExport.Click lo usaremos más adelante para exportar el DataTable a Excel

//            // Orden de controles en la pestaña de informes:
//            // 1) Grid (Fill)
//            // 2) Lista de conexiones (Left)
//            // 3) Panel superior con botones (Top)
//            tabPredefinidos.Controls.Add(_grid);           // Dock.Fill
//            tabPredefinidos.Controls.Add(_chkConnections); // Dock.Left
//            tabPredefinidos.Controls.Add(_topPanel);       // Dock.Top

//            _tabControl.TabPages.Add(tabPredefinidos);
//            _tabControl.TabPages.Add(tabAdHoc);

//            Controls.Add(_tabControl);
//        }

//        private void CargarConexiones()
//        {
//            var conexiones = _connectionCatalog
//                .GetAllConnections()
//                .OrderBy(c => c.Type)
//                .ThenBy(c => c.Id);

//            _chkConnections.Items.Clear();

//            // Ninguna marcada por defecto
//            foreach (ConnectionInfo c in conexiones)
//            {
//                _chkConnections.Items.Add(c, false);
//            }
//        }

//        private void SetAllConnectionsChecked(bool isChecked)
//        {
//            for (int i = 0; i < _chkConnections.Items.Count; i++)
//            {
//                _chkConnections.SetItemChecked(i, isChecked);
//            }
//        }

//        private string[] GetSelectedConnectionIds()
//        {
//            return _chkConnections.CheckedItems
//                .OfType<ConnectionInfo>()
//                .Select(c => c.Id)
//                .ToArray();
//        }

//        private async Task ProbarConsultaAsync()
//        {
//            var selectedIds = GetSelectedConnectionIds();

//            if (selectedIds.Length == 0)
//            {
//                MessageBox.Show(
//                    "Selecciona al menos una conexión para probar.",
//                    "Sin selección",
//                    MessageBoxButtons.OK,
//                    MessageBoxIcon.Warning);

//                return;
//            }

//            try
//            {
//                const string sql = "SELECT SYSDATE AS FECHA_ACTUAL FROM DUAL";

//                DataTable? combined = null;

//                foreach (var connectionId in selectedIds)
//                {
//                    // Ejecutamos la consulta en cada conexión
//                    var table = await _queryExecutor.ExecuteQueryAsync(
//                        sql,
//                        new Dictionary<string, object?>(),
//                        connectionId);

//                    if (combined is null)
//                    {
//                        // Creamos el DataTable combinado con las columnas del primero
//                        combined = table.Clone();
//                        // Añadimos una columna para saber de qué conexión viene cada fila
//                        combined.Columns.Add("CONEXION_ID", typeof(string));
//                    }

//                    foreach (DataRow row in table.Rows)
//                    {
//                        var newRow = combined.NewRow();

//                        foreach (DataColumn col in table.Columns)
//                        {
//                            newRow[col.ColumnName] = row[col];
//                        }

//                        newRow["CONEXION_ID"] = connectionId;
//                        combined.Rows.Add(newRow);
//                    }
//                }

//                _grid.DataSource = combined;
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(
//                    $"Error ejecutando la consulta de prueba:{Environment.NewLine}{ex.Message}",
//                    "Error en consulta",
//                    MessageBoxButtons.OK,
//                    MessageBoxIcon.Error);
//            }
//        }

//    }
//}




using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Infrastructure.Data;
using OracleReportExport.Infrastructure.Interfaces;
using OracleReportExport.Infrastructure.Services;

namespace OracleReportExport.Presentation.Desktop
{
    public class MainForm : Form
    {
        private readonly TabControl _tabControl = new();

        private readonly DataGridView _grid = new()
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false
        };

        // Lista de conexiones (Central + estaciones)
        private readonly CheckedListBox _chkConnections = new()
        {
            Dock = DockStyle.Left,
            Width = 260,
            CheckOnClick = true
        };

        // Botones de acciones (con AutoSize para que se vea bien el texto)
        private readonly Button _btnSelectAll = new()
        {
            Text = "Marcar todas",
            AutoSize = true
        };

        private readonly Button _btnUnselectAll = new()
        {
            Text = "Desmarcar",
            AutoSize = true
        };

        private readonly Button _btnExport = new()
        {
            Text = "Exportar a Excel",
            AutoSize = true
        };

        private readonly Button _btnTest = new()
        {
            Text = "Probar informe real",
            AutoSize = true
        };

        // Panel superior para los botones
        private readonly FlowLayoutPanel _topPanel = new()
        {
            Dock = DockStyle.Top,
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };

        // Servicios
        private readonly ConnectionCatalogService _connectionCatalog;
        private readonly IReportService _reportService;

        public MainForm()
        {
            Text = "Oracle Report Export";
            Width = 1200;
            Height = 800;

            _tabControl.Dock = DockStyle.Fill;

            var tabPredefinidos = new TabPage("Informes predefinidos");
            var tabAdHoc = new TabPage("SQL avanzada");

            // Inicializar servicios
            _connectionCatalog = new ConnectionCatalogService();

            IOracleConnectionFactory connectionFactory = new OracleConnectionFactory();
            IQueryExecutor queryExecutor = new OracleQueryExecutor(connectionFactory);
            var reportDefinitionRepository = new JsonReportDefinitionRepository();
            _reportService = new ReportService(reportDefinitionRepository, queryExecutor);

            // Cargar conexiones en el CheckedListBox
            CargarConexiones();

            // Configurar barra superior de botones
            _topPanel.Controls.Add(_btnSelectAll);
            _topPanel.Controls.Add(_btnUnselectAll);
            _topPanel.Controls.Add(_btnExport);
            _topPanel.Controls.Add(_btnTest);

            // Eventos de botones
            _btnSelectAll.Click += (_, __) => SetAllConnectionsChecked(true);
            _btnUnselectAll.Click += (_, __) => SetAllConnectionsChecked(false);
            _btnTest.Click += async (_, __) => await ProbarInformeRealAsync();
            // _btnExport.Click lo usaremos más adelante para exportar el DataTable a Excel

            // Orden de controles en la pestaña de informes:
            // 1) Grid (Fill)
            // 2) Lista de conexiones (Left)
            // 3) Panel superior con botones (Top)
            tabPredefinidos.Controls.Add(_grid);           // Dock.Fill
            tabPredefinidos.Controls.Add(_chkConnections); // Dock.Left
            tabPredefinidos.Controls.Add(_topPanel);       // Dock.Top

            _tabControl.TabPages.Add(tabPredefinidos);
            _tabControl.TabPages.Add(tabAdHoc);

            Controls.Add(_tabControl);
        }

        private void CargarConexiones()
        {
            var conexiones = _connectionCatalog
                .GetAllConnections()
                .OrderBy(c => c.Type)
                .ThenBy(c => c.Id);

            _chkConnections.Items.Clear();

            // Ninguna marcada por defecto
            foreach (ConnectionInfo c in conexiones)
            {
                _chkConnections.Items.Add(c, false);
            }
        }

        private void SetAllConnectionsChecked(bool isChecked)
        {
            for (int i = 0; i < _chkConnections.Items.Count; i++)
            {
                _chkConnections.SetItemChecked(i, isChecked);
            }
        }

        private string[] GetSelectedConnectionIds()
        {
            return _chkConnections.CheckedItems
                .OfType<ConnectionInfo>()
                .Select(c => c.Id)
                .ToArray();
        }

        private async Task ProbarInformeRealAsync()
        {
            var selectedIds = GetSelectedConnectionIds();

            if (selectedIds.Length == 0)
            {
                MessageBox.Show(
                    "Selecciona al menos una conexión para ejecutar el informe.",
                    "Sin selección",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            try
            {
                // Informe real definido en Reports.json
                //const string reportId = "listado_motocicletas";

                //var parametros = new Dictionary<string, object?>(); // sin parámetros 

                ////////var parametros = new Dictionary<string, object?>
                ////////{
                ////////    // Fechas
                ////////    //listado_motocicletas
                ////////    ["FechaDesde"] = new DateTime(2025, 10, 6, 0, 0, 0),
                ////////    ["FechaHasta"] = new DateTime(2025, 11, 14, 23, 59, 59),

                ////////    // Tipos de vehículo
                ////////    ["TipoVehiculo1"] = "CI",
                ////////    ["TipoVehiculo2"] = "M",

                ////////    // Anulada
                ////////    ["Anulada"] = 0,   // 0 = no anulada

                ////////    // Categorías (LIKE)
                ////////    ["Cat1"] = "%L1%",
                ////////    ["Cat2"] = "%L1e%",
                ////////    ["Cat3"] = "%L3%",
                ////////    ["Cat4"] = "%L3e%",
                ////////    ["Cat5"] = "%L4e%"
                ////////};

                const string reportId = "Enac_NoPeriodicas";    
                var parametros = new Dictionary<string, object?>
                {
                    ["Anulada"] = 0,
                    ["FechaDesde"] = new DateTime(2025, 1, 1, 0, 0, 0),
                    ["FechaHasta"] = new DateTime(2025, 10, 15, 23, 59, 59),
                    ["EsPeriodicaPura"] = "N",

                    ["CodTipo1"] = "RC",
                    ["CodTipo2"] = "RS",
                    ["CodTipo3"] = "PU",
                    ["CodTipo4"] = "PM",
                    ["CodTipo5"] = "DP",
                    ["CodTipo6"] = "CS",
                    ["CodTipo7"] = "OT"
                };

                var table = await _reportService.ExecuteReportAsync(
                    reportId,
                    parametros,
                    selectedIds);

                _grid.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error ejecutando el informe:{Environment.NewLine}{ex.Message}",
                    "Error en informe",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}


