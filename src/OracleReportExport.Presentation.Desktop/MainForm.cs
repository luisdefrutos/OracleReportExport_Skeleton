using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Models;
using OracleReportExport.Infrastructure.Configuration;
using OracleReportExport.Infrastructure.Data;
using OracleReportExport.Infrastructure.Interfaces;
using OracleReportExport.Infrastructure.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OracleReportExport.Presentation.Desktop
{
    public class MainForm : Form
    {
        private readonly TabControl _tabControl = new();


        private readonly DataGridView _grid = new()
        {
             Dock = DockStyle.Fill,
             ReadOnly = true,
             AllowUserToAddRows = false,
             AllowUserToResizeColumns = false,
             AllowUserToResizeRows = false,
             ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
             RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            ScrollBars = ScrollBars.Both
        };

        // Lista de conexiones (Central + estaciones)
        private readonly CheckedListBox _chkConnections = new()
        {
            Dock = DockStyle.Left,
            Width = 260,
            CheckOnClick = true
        };

        // Botones de acciones de la barra superior
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

        private readonly Button _btnRunReport = new()
        {
            Text = "Ejecutar informe",
            AutoSize = true
        };

        // Combo de informes
        private readonly ComboBox _cmbReports = new()
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 260
        };

        // Panel superior
        private readonly FlowLayoutPanel _topPanel = new()
        {
            Dock = DockStyle.Top,
            Height = 42,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(5, 5, 5, 0)
        };

        // Grupo de parámetros
        private readonly GroupBox _grpParametros = new()
        {
            Text = "Parámetros",
            Dock = DockStyle.Top,
            Height = 120,
            Padding = new Padding(8, 18, 8, 8)
        };

        // Tabla donde se colocan dinámicamente los parámetros (Label + Control)
        private readonly TableLayoutPanel _paramsTable = new()
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            Padding = new Padding(4),
        };

        // Mapa NombreParametro -> Control generado
        private readonly Dictionary<string, Control> _parameterControls = new();

        // Servicios
        private readonly ConnectionCatalogService _connectionCatalog;
        private readonly IReportService _reportService;
        private readonly IReportDefinitionRepository _reportDefinitionRepository;
        private readonly IQueryExecutor _queryExecutor;
        private readonly IOracleConnectionFactory _connectionFactory;

        // Informe actual
        private ReportDefinition? _currentReport;

        public MainForm()
        {
            Text = "Oracle Report Export";

            WindowState = FormWindowState.Maximized;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedSingle;  // o FixedDialog

            _tabControl.Dock = DockStyle.Fill;

            var tabPredefinidos = new TabPage("Informes predefinidos");
            var tabAdHoc = new TabPage("SQL avanzada");

            // Inicializar servicios
            _connectionCatalog = new ConnectionCatalogService();

            _connectionFactory = new OracleConnectionFactory();
            _queryExecutor = new OracleQueryExecutor(_connectionFactory);

            _reportDefinitionRepository = new JsonReportDefinitionRepository();
            _reportService = new ReportService(_reportDefinitionRepository, _queryExecutor);



            // Cargar conexiones
            CargarConexiones();

            // Configurar diseño
            ConfigureTopPanel();
            ConfigureParametrosGroup();

            // Orden en la pestaña de informes:
            // 1) Grid (Fill)
            // 2) Lista conexiones (Left)
            // 3) Grupo de parámetros (Top)
            // 4) Panel botones (Top)
            tabPredefinidos.Controls.Add(_grid);
            tabPredefinidos.Controls.Add(_chkConnections);
            tabPredefinidos.Controls.Add(_grpParametros);
            tabPredefinidos.Controls.Add(_topPanel);

            _tabControl.TabPages.Add(tabPredefinidos);
            _tabControl.TabPages.Add(tabAdHoc);

            Controls.Add(_tabControl);

            // Al cargar el formulario, rellenar combo de informes
            Load += async (_, __) => await LoadReportsAsync();
        }

        #region Diseño


  
        private void ConfigureTopPanel()
        {
            _topPanel.Controls.Add(_btnSelectAll);
            _topPanel.Controls.Add(_btnUnselectAll);
            _topPanel.Controls.Add(_btnExport);

            var sep = new Label
            {
                AutoSize = true,
                Margin = new Padding(20, 10, 0, 0),
                Text = "|"
            };
            _topPanel.Controls.Add(sep);

            var lblInforme = new Label
            {
                Text = "Informe:",
                AutoSize = true,
                Margin = new Padding(20, 10, 0, 0)
            };
            _topPanel.Controls.Add(lblInforme);

            _topPanel.Controls.Add(_cmbReports);
            _topPanel.Controls.Add(_btnRunReport);

            // Eventos
            _btnSelectAll.Click += (_, __) => SetAllConnectionsChecked(true);
            _btnUnselectAll.Click += (_, __) => SetAllConnectionsChecked(false);
            // Export lo dejamos para más adelante
            _btnRunReport.Click += BtnRunReport_Click;
            _cmbReports.SelectedIndexChanged += CmbReports_SelectedIndexChanged;
        }
        private void SetAllConnectionsChecked(bool isChecked)
        {
            for (int i = 0; i < _chkConnections.Items.Count; i++)
            {
                _chkConnections.SetItemChecked(i, isChecked);
            }
        }

        private void ConfigureParametrosGroup()
        {
            _grpParametros.Controls.Add(_paramsTable);
            _paramsTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));   // Label
            _paramsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // Control
        }

        #endregion

        #region Carga de datos inicial

        private void CargarConexiones()
        {
            var conexiones = _connectionCatalog
                .GetAllConnections()
                .OrderBy(c => c.Type)
                .ThenBy(c => c.Id);

            _chkConnections.Items.Clear();

            foreach (ConnectionInfo c in conexiones)
            {
                _chkConnections.Items.Add(c, false);
            }
        }

        private async Task LoadReportsAsync()
        {
            var reports = (await _reportService.GetAvailableReportsAsync()).ToList();

            _cmbReports.DataSource = reports;
            _cmbReports.DisplayMember = nameof(ReportDefinition.Name);
            _cmbReports.ValueMember = nameof(ReportDefinition.Id);

            if (reports.Any())
            {
                _cmbReports.SelectedIndex = 0;
            }
        }

        #endregion

        #region Parámetros dinámicos

        private void CmbReports_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_cmbReports.SelectedItem is ReportDefinition report)
            {
                CargarConexiones();
                this._grid.DataSource = null;
                _currentReport = report;
                RenderParameters(report);
            }
        }

        private void RenderParameters(ReportDefinition report)
        {
            _parameterControls.Clear();

            _paramsTable.SuspendLayout();

            // Limpiamos todo UNA sola vez
            _paramsTable.Controls.Clear();
            _paramsTable.RowStyles.Clear();
            _paramsTable.ColumnStyles.Clear();
            _paramsTable.ColumnCount = 2;
            _paramsTable.RowCount = 0;

            // Caso: sin ningún tipo de parámetro
            bool hasMaster = report.TableMasterForParameters != null && report.TableMasterForParameters.Count > 0;
            bool hasParams = report.Parameters != null && report.Parameters.Count > 0;

            if (!hasMaster && !hasParams)
            {
                _grpParametros.Height = 70;

                var lbl = new Label
                {
                    Text = "Este informe no requiere parámetros.",
                    AutoSize = true,
                    ForeColor = SystemColors.GrayText,
                    Margin = new Padding(4, 8, 4, 4)
                };

                _paramsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                _paramsTable.RowCount = 1;
                _paramsTable.Controls.Add(lbl, 0, 0);
                _paramsTable.SetColumnSpan(lbl, 2);

                _paramsTable.ResumeLayout();
                return;
            }


            int row = 0;

            //  Primero pintamos los TableMasterForParameters (combos)
            if (hasMaster)
            {
                foreach (var p in report.TableMasterForParameters!)
                {
                    _paramsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    _paramsTable.RowCount = row + 1;

                    var label = new Label
                    {
                        Text = p.Label ?? string.Empty,
                        AutoSize = true,
                        Margin = new Padding(4, 6, 4, 4)
                    };

                    Control? input = CreateControlForTableMasterParameter(p);

                    if (input != null)
                    {
                        _paramsTable.Controls.Add(label, 0, row);
                        _paramsTable.Controls.Add(input, 1, row);

                        _parameterControls[p.Name] = input;
                        row++;
                    }
                }
            }

            // Luego pintamos los Parameters “normales” (fechas, anulada, etc.)
            if (hasParams)
            {
                foreach (var p in report.Parameters!)
                {
                    _paramsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    _paramsTable.RowCount = row + 1;

                    var label = new Label
                    {
                        Text = p.Label ?? p.Name,
                        AutoSize = true,
                        Margin = new Padding(4, 6, 4, 4)
                    };

                    Control? input = CreateControlForParameter(p);

                    if (input != null)
                    {
                        _paramsTable.Controls.Add(label, 0, row);
                        _paramsTable.Controls.Add(input, 1, row);

                        _parameterControls[p.Name] = input;
                        row++;
                    }
                }
            }
            _paramsTable.ResumeLayout();
            _grpParametros.Height =
                _paramsTable.Controls
                            .Cast<Control>()
                            .Sum(c => c.Height)
                + 15;

        }


        private Control? CreateControlForParameter(ReportParameterDefinition parameter)
        {
            var type = (parameter.Type ?? "string").ToLowerInvariant();

            switch (type)
            {
                case "date":
                    return new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Short,
                        Width = 120,
                        Margin = new Padding(4, 2, 4, 2)
                    };

                case "bool":
                    return new CheckBox
                    {
                        Checked = parameter.IsRequired?true:false,
                        AutoSize = true,
                        Margin = new Padding(4, 4, 4, 2)
                    };
                case "funcion":
                    return new CheckBox
                    {
                        Checked = parameter.IsRequired ? true : false,
                        AutoSize = true,
                        Margin = new Padding(4, 4, 4, 2)
                    };

                default:  
                    return null;
            }
        }


        private Control? CreateControlForTableMasterParameter(TableMasterParameterDefinition parameter)
        {
            var type = (parameter.Type ?? "string").ToLowerInvariant();

            switch (type)
            {
                case "combobox":
                    var initialConnection = _connectionCatalog.GetAllConnections().Where(x => x.Type.Contains("Estacion")).FirstOrDefault();
                    return LoadTableMasterDataIntoControl(parameter);
                default:
                    return null;
            }
        }

        private Control LoadTableMasterDataIntoControl(TableMasterParameterDefinition parameter)
        {
            var initialConnection = _connectionCatalog
                .GetAllConnections()
                .FirstOrDefault(x => x.Type.Contains("Estacion"));

            if (initialConnection == null)
                throw new Exception("No se encontró una conexión válida para Estación.");

            if (string.IsNullOrWhiteSpace(parameter.SqlQueryMaster))
                return new CheckedListBox();   // vacío

            DataTable dt = new DataTable();
            using (var conn = _connectionFactory.CreateConnection(initialConnection.Id) as OracleConnection)
            {
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = parameter.SqlQueryMaster;
                using var da = new OracleDataAdapter(cmd);
                var ds = new DataSet();
                da.Fill(ds);
                if (ds.Tables.Count == 0)
                    return new CheckedListBox();
                dt = ds.Tables[0];
            }

            var clb = new CheckedListBox
            {
                CheckOnClick = true,
                IntegralHeight = false,
                Height = Math.Min(150, 18 * dt.Rows.Count + 4)   // alto razonable
            };

            clb.DisplayMember = parameter.Text ?? String.Empty;
            clb.ValueMember = parameter.Id ?? String.Empty;
            // Valores que deben aparecer ya marcados (CI, M, etc.)
            var preselected = parameter.ValuesRequired ?? new List<string>();

            int maxWidth = 0;
            foreach (DataRow row in dt.Rows)
            {
                string value = row[parameter.Id]?.ToString() ?? "";
                string text = row[parameter.Text]?.ToString() ?? "";
                // objeto para usar DisplayMember/ValueMember
                var item = new MultiItem
                {
                    Value = value,
                    Text = text
                };

                int index = clb.Items.Add(item);
                if (preselected.Contains(value))
                    clb.SetItemChecked(index, true);

                int w = TextRenderer.MeasureText(text, clb.Font).Width;
                if (w > maxWidth)
                    maxWidth = w;

            }
            clb.Height = Math.Min(150, 18 * clb.Items.Count + 4);
            clb.Width = maxWidth + SystemInformation.VerticalScrollBarWidth + 30;
            return clb;
        }


        private Dictionary<string, object?> BuildParametersFromUI()
        {
            var result = new Dictionary<string, object?>();
            if (_currentReport == null)
                return result;

            //
            // Parámetros “normales” (FechaDesde, FechaHasta, Anulada, etc.)

            //
            if (_currentReport.Parameters != null && _currentReport.Parameters.Count > 0)
            {
                foreach (var p in _currentReport.Parameters)
                {
                  if (!_parameterControls.TryGetValue(p.Name, out var ctrl))
                        continue;

                    object? value = ctrl switch
                    {
                        DateTimePicker dtp => dtp.Value,
                        NumericUpDown nud => Convert.ToInt32(nud.Value),
                        CheckBox chk => chk.Checked,
                        TextBox txt => string.IsNullOrWhiteSpace(txt.Text) ? null : txt.Text,
                        _ => null
                    };

                    if (p.IsRequired &&
                        (value == null || (value is string s && string.IsNullOrWhiteSpace(s))))
                    {
                        MessageBox.Show(
                            $"El parámetro \"{p.Label ?? p.Name}\" es obligatorio.",
                            "Parámetros incompletos",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        return new Dictionary<string, object?>();
                    }
                    if (value != null && value.GetType().Name.ToUpper().Equals("Boolean".ToUpper()))
                    {
                        if (p.Type == "funcion")
                        {
                            int numberFromBoolean = value is bool b ? (b ? 1 : 0) : 0;
                            value = p.Values.Where(x => x.Key == numberFromBoolean).FirstOrDefault()?.Value ?? String.Empty;
                        }
                        else
                            value = (bool)value ? -1 : 0;
                    }
                    switch (p.Name.ToUpper())
                    {
                        case "FECHADESDE":
                             var fromDate = (DateTime)value;
                            var fromDateFormat = new DateTime(fromDate.Year, fromDate.Month, fromDate.Day);
                            value = fromDateFormat;
                            break;
                            case "FECHAHASTA":
                            var toDate = (DateTime)value;
                            var toDateFormat = new DateTime(toDate.Year, toDate.Month, toDate.Day, 23, 59, 59);
                            value = toDateFormat;
                            break;
                        default:
                            break;
                    }                     
                    result[p.Name] = value;
                }
            }

            //
            //Multi-select: construir las listas para IN(...)
            //    TIPOSVEHICULO  -> {TiposVehiculoList}
            //    TIPOINSPECCION_ENAC -> {CategoriasList}
            //

            List<string> GetCheckedCodes(string controlKey)
            {
                if (_parameterControls.TryGetValue(controlKey, out var ctrl) &&
                    ctrl is CheckedListBox clb)
                {
                    return clb.CheckedItems
                              .Cast<MultiItem>()
                              .Select(i => i.Value)
                              .ToList();
                }

                return new List<string>();
            }

            string BuildInList(List<string> values)
            {
                if (values == null || values.Count == 0)
                    return "''";  // si quieres obligar a elegir algo, aquí podrías lanzar excepción

                return string.Join(", ",
                    values.Select(v => $"'{v.Replace("'", "''")}'"));
            }
            //TODO:Cambiar los nombres de los parámetros según convenga y archivo txt de informes (campos interpolados)
            // TIPOSVEHICULO 
            var tiposVehiculo = GetCheckedCodes("TIPOSVEHICULO");
             if(tiposVehiculo!=null && tiposVehiculo.Count>0)
              result["TiposVehiculoList"] = BuildInList(tiposVehiculo);
            //TODO:Cambiar los nombres de los parámetros según convenga y archivo txt de informes (campos interpolados)
            // CATEGORIAS
            var categorias = GetCheckedCodes("CATEGORIAS");
            if (categorias != null && categorias.Count > 0)
                result["CategoriasList"] = BuildInList(categorias);

            //
            //   Validación final de requeridos (por si mantienes TipoVehiculo1.. y Cat1.. en el JSON)
            //    -> los excluimos de la validación para que no den guerra
            //
            if (_currentReport.Parameters != null)
            {
                foreach (var p in _currentReport.Parameters.Where(x => x.IsRequired))
                {
                      if (!result.TryGetValue(p.Name, out var val) ||
                        val == null ||
                        (val is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        MessageBox.Show(
                            $"El parámetro \"{p.Label ?? p.Name}\" es obligatorio.",
                            "Parámetros incompletos",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        return new Dictionary<string, object?>();
                    }
                }
            }

            return result;
        }

        #endregion

        #region Ejecución de informes
        private List<ConnectionInfo> GetSelectedConnection()
        {
            return _chkConnections.CheckedItems
                .OfType<ConnectionInfo>().ToList();
                 
        }

        private async void BtnRunReport_Click(object? sender, EventArgs e)
        {
            var listConnectionsActive = GetSelectedConnection();

            if (listConnectionsActive==null || listConnectionsActive.Count == 0)
            {
                MessageBox.Show(
                    "Selecciona al menos una conexión para ejecutar el informe.",
                    "Sin selección",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (_cmbReports.SelectedItem is not ReportDefinition report)
            {
                MessageBox.Show(
                    "Selecciona un informe en el combo.",
                    "Informe no seleccionado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _currentReport = report;

            var parametros = BuildParametersFromUI();
            // Si la validación devolvió un diccionario vacío porque faltan requeridos, salimos.
            if (_currentReport.Parameters != null &&
                _currentReport.Parameters.Count > 0 &&
                parametros.Count == 0)
            {
                return;
            }

            using var loading = new LoadingForm("Cargando datos...");

            try
            {
                loading.Owner = this; 
                loading.Show();
                loading.Refresh();

                Enabled = false;
                Cursor = Cursors.WaitCursor;

                var table = await _reportService.ExecuteReportAsync(
                    report,
                    parametros,
                    listConnectionsActive);

                _grid.DataSource = table;
            }
            catch (OracleException ex) when (ex.Number == 942) // ORA-00942
            {
                MessageBox.Show(
                    "La tabla o vista no existe en la base de datos.Verifique que esta ejecutando " +
                    "el informe correcto en la base de datos correspondiente.\n\nEsta ejecutando " +
                    "un informe de "+ _currentReport.SourceType.ToString()+ ".\n",
                    "Error de base de datos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (OracleException ex)
            {
                MessageBox.Show(
                    $"Error de Oracle: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error ejecutando el informe:{Environment.NewLine}{ex.Message}",
                    "Error en informe",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                loading.Close();
                Enabled = true;
            }
        }

        #endregion

        #region Form de carga
        private sealed class LoadingForm : Form
        {
            public LoadingForm(string message)
            {
                // Lo vamos a centrar nosotros → Manual
                StartPosition = FormStartPosition.Manual;
                TopMost = true;
                ShowInTaskbar = false;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                ControlBox = false;
                Width = 260;
                Height = 100;
                Text = string.Empty;

                var lbl = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Text = message,
                    Font = new Font(SystemFonts.DefaultFont.FontFamily, 10, FontStyle.Bold)
                };

                Controls.Add(lbl);
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);

                // Si tiene formulario padre, centramos sobre él
                if (Owner != null)
                {
                    var rect = Owner.Bounds;
                    Left = rect.Left + (rect.Width - Width) / 2;
                    Top = rect.Top + (rect.Height - Height) / 2;
                }
                else
                {
                    // Si no tiene, centramos en la pantalla activa
                    var screen = Screen.FromPoint(Cursor.Position).WorkingArea;
                    Left = screen.Left + (screen.Width - Width) / 2;
                    Top = screen.Top + (screen.Height - Height) / 2;
                }
            }
        }

        #endregion
    }

    public class MultiItem
    {
        public string Value { get; set; } = "";
        public string Text { get; set; } = "";
       public override string ToString()
            => $"({Value}) -> {Text}";
    }

}



