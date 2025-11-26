using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using Oracle.ManagedDataAccess.Client;
using OracleReportExport.Application.Interfaces;
using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Enums;
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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace OracleReportExport.Presentation.Desktop
{
    public class MainForm : Form
    {
        #region Campos privados

        private readonly TabControl _tabControl = new();
        private TabPage _tabPredefinidos;


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

        private readonly CheckedListBox _chkConnections = new()
        {
            Dock = DockStyle.Left,
            Width = 260,
            CheckOnClick = true,
            ScrollAlwaysVisible = true
        };

        private readonly CheckedListBox _chkConnectionsAdHoc = new()
        {
            Dock = DockStyle.Left,
            Width = 260,
            CheckOnClick = true,
            ScrollAlwaysVisible = true
        };

        private readonly Button _btnSelectAll = new()
        {
            Text = "Marcar todas",
            AutoSize = true,
            Visible = true,
            Tag = ResultTabUI.TabInitial
        };

        private readonly Button _btnUnselectAll = new()
        {
            Text = "Desmarcar",
            AutoSize = true,
            Visible = true,
            Tag = ResultTabUI.TabInitial
        };

        private readonly Button _btnSelectAllAdHoc = new()
        {
            Text = "Marcar todas",
            AutoSize = true,
            Visible = true,
            Tag = ResultTabUI.TabSecundary
        };

        private readonly Button _btnUnselectAllAdHoc = new()
        {
            Text = "Desmarcar",
            AutoSize = true,
            Visible = true,
            Tag = ResultTabUI.TabSecundary
        };

        private readonly Button _btnClearAdHoc = new()
        {
            Text = "Limpiar consulta personalizada",
            AutoSize = true,
            Visible = true,
            Tag = ResultTabUI.TabSecundary
        };

        private readonly Button _btnExport = new()
        {
            Text = "Exportar a Excel",
            AutoSize = true,
            Visible = true
        };

        private readonly Button _btnRunReport = new()
        {
            Text = "Ejecutar informe",
            AutoSize = true,
            Visible = true
        };

        private readonly Button _btnVerConsulta = new()
        {
            Text = "Ver Consulta",
            AutoSize = true,
            Visible = true
        };

        private readonly ComboBox _cmbReports = new()
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 260
        };

        private readonly FlowLayoutPanel _topPanel = new()
        {
            Dock = DockStyle.Top,
            Height = 42,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(5, 5, 5, 0)
        };

        private readonly FlowLayoutPanel _topPanelAdHoc = new()
        {
            Dock = DockStyle.Top,
            Height = 52,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(5, 5, 5, 0)
        };

        private readonly GroupBox _grpParametros = new()
        {
            Text = "Parámetros",
            Dock = DockStyle.Top,
            Height = 120,
            Padding = new Padding(8, 18, 8, 8)
        };

        private readonly FlowLayoutPanel _paramsPanel = new()
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(8),
            Margin = new Padding(0)
        };

        private readonly RichTextBox _txtSqlAdHoc = new()
        {
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 10f),
            BackColor = Color.FromArgb(232, 238, 247),   // (#E8EEF7)
            ForeColor = Color.FromArgb(28, 59, 106),     // (#1C3B6A)
            ScrollBars = RichTextBoxScrollBars.Both,
            BorderStyle = BorderStyle.FixedSingle,
            ShortcutsEnabled = true
        };

        private readonly Button ButtonAdHoc = new()
        {
            AutoSize = true,
            Text = "Ejecutar SQL",
            Visible = true,

        };


        // Grid para resultados de SQL avanzada
        private readonly DataGridView _gridAdHoc = new()
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

        private readonly Dictionary<string, Control> _parameterControls = new();

        private Label? _lblCountRows;
        private Button? _btnExcel;
        private Label? _lblCountRowsAdHoc;
        private Button? _btnExcelAdHoc;

        private readonly ConnectionCatalogService _connectionCatalog;
        private readonly IReportService _reportService;
        private readonly IReportDefinitionRepository _reportDefinitionRepository;
        private readonly IQueryExecutor _queryExecutor;
        private readonly IOracleConnectionFactory _connectionFactory;

        private ReportDefinition? _currentReport;

        #endregion

        #region Constructor y carga inicial

        public MainForm()
        {
            Text = "Oracle Report Export";
            WindowState = FormWindowState.Maximized;
            MaximizeBox = true;
            FormBorderStyle = FormBorderStyle.Sizable;

            _tabControl.Dock = DockStyle.Fill;

            _tabPredefinidos = new TabPage("Informes predefinidos");
            var tabAdHoc = new TabPage("SQL avanzada");

            _connectionCatalog = new ConnectionCatalogService();
            _connectionFactory = new OracleConnectionFactory();
            _queryExecutor = new OracleQueryExecutor(_connectionFactory);
            _reportDefinitionRepository = new JsonReportDefinitionRepository();
            _reportService = new ReportService(_reportDefinitionRepository, _queryExecutor);

            CargarConexiones();
            ConfigurarTopPanel();
            ConfigurarGrupoParametros();

            // --- Pestaña de informes predefinidos ---
            _tabPredefinidos.Controls.Add(_grid);
            _tabPredefinidos.Controls.Add(_chkConnections);
            _tabPredefinidos.Controls.Add(_grpParametros);
            _tabPredefinidos.Controls.Add(_topPanel);

            // --- Pestaña SQL avanzada ---

            // Copiamos las mismas conexiones al modo ad-hoc
            _chkConnectionsAdHoc.Items.AddRange(
                _chkConnections.Items.OfType<ConnectionInfo>().ToArray()
            );

            // Panel superior de título
            _topPanelAdHoc.Controls.Add(new Label
            {
                Text = "Creación de Consultas Personalizadas",
                AutoSize = true,
                Padding = new Padding(8, 10, 8, 8)
            });
            var sepAdHoc = new Label
            {
                AutoSize = true,
                Margin = new Padding(20, 10, 10, 0),
                Text = "|"
            };
            _topPanelAdHoc.Controls.Add(sepAdHoc);

            _btnSelectAllAdHoc.Anchor = AnchorStyles.Left;
            _btnSelectAllAdHoc.Margin = new Padding(0, 5, 10, 5);
            _btnSelectAllAdHoc.Tag = ResultTabUI.TabSecundary;
            _btnSelectAllAdHoc.Click += _btnSelectAllAdHoc_Click;
            _topPanelAdHoc.Controls.Add(_btnSelectAllAdHoc);


            _btnUnselectAllAdHoc.Anchor = AnchorStyles.Left;
            _btnUnselectAllAdHoc.Margin = new Padding(0, 5, 10, 5);
            _btnUnselectAllAdHoc.Tag = ResultTabUI.TabSecundary;
            _btnUnselectAllAdHoc.Click += _btnUnselectAllAdHoc_Click;
            _topPanelAdHoc.Controls.Add(_btnUnselectAllAdHoc);



            _btnClearAdHoc.Anchor = AnchorStyles.Left;
            _btnClearAdHoc.Margin = new Padding(0, 5, 10, 5);
            _btnClearAdHoc.Tag = ResultTabUI.TabSecundary;
            _btnClearAdHoc.Click += _btnClearAdHoc_Click;
            _topPanelAdHoc.Controls.Add(_btnClearAdHoc);


            ButtonAdHoc.Anchor = AnchorStyles.Left;
            ButtonAdHoc.Margin = new Padding(0, 5, 10, 5);
            ButtonAdHoc.Click += ButtonAdHoc_Click;
            _topPanelAdHoc.Controls.Add(ButtonAdHoc);




            // Panel derecho que contiene editor + botón + grid
            var rightPanelAdHoc = new Panel
            {
                Dock = DockStyle.Fill
            };

            // Layout vertical: SQL (50%), botón (auto), grid (50%)
            var layoutAdHoc = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3
            };

            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));   // RichTextBox
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // Botón
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));   // Grid

            // Fila 0: editor SQL
            _txtSqlAdHoc.Dock = DockStyle.Fill;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Copiar", null, (_, __) => _txtSqlAdHoc.Copy());
            menu.Items.Add("Pegar", null, (_, __) => _txtSqlAdHoc.Paste());
            menu.Items.Add("Cortar", null, (_, __) => _txtSqlAdHoc.Cut());
            menu.Items.Add("Seleccionar todo", null, (_, __) => _txtSqlAdHoc.SelectAll());

            _txtSqlAdHoc.ContextMenuStrip = menu;

            layoutAdHoc.Controls.Add(_txtSqlAdHoc, 0, 0);
            // Fila 2: grid resultados
            _gridAdHoc.Dock = DockStyle.Fill;
            layoutAdHoc.Controls.Add(_gridAdHoc, 0, 2);
            rightPanelAdHoc.Controls.Add(layoutAdHoc);
            // Orden en la pestaña:
            // 1) panel derecho (Fill)
            // 2) conexiones (Left)
            // 3) cabecera (Top)
            tabAdHoc.Controls.Add(rightPanelAdHoc);
            tabAdHoc.Controls.Add(_chkConnectionsAdHoc);
            tabAdHoc.Controls.Add(_topPanelAdHoc);

            // Añadimos pestañas al TabControl
            _tabControl.TabPages.Add(_tabPredefinidos);
            _tabControl.TabPages.Add(tabAdHoc);

            Controls.Add(_tabControl);

            Load += MainForm_LoadAsync;
        }

        private void _btnClearAdHoc_Click(object? sender, EventArgs e)
        {
            if (_txtSqlAdHoc != null)
                _txtSqlAdHoc.Clear();
        }

        private void _btnUnselectAllAdHoc_Click(object? sender, EventArgs e)
        {
            if (sender != null && sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(false, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(false, _chkConnectionsAdHoc);
                        break;
                    default:
                        break;
                }
            }
        }

        private void _btnSelectAllAdHoc_Click(object? sender, EventArgs e)
        {
            if (sender != null && sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(true, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(true, _chkConnectionsAdHoc);
                        break;
                    default:
                        break;
                }
            }
        }

        private async void ButtonAdHoc_Click(object? sender, EventArgs e)
        {
            var listConnectionsActive = GetSelectedConnectionsAdHoc();

            if (listConnectionsActive == null || listConnectionsActive.Count == 0)
            {
                MessageBox.Show(
                    "Selecciona al menos una conexión para ejecutar el informe.",
                    "Sin selección",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            using var cts = new CancellationTokenSource();
            using (var loadingFormAdHoc = new LoadingForm("Cargando Datos Consulta ....", cts))
            {
                try
                {
                     RemoveControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary);
                    RecursiveEnableControlsForm(this, false);
                    loadingFormAdHoc.Owner = this;
                    loadingFormAdHoc.Show();
                    loadingFormAdHoc.Refresh();
                    Enabled = false;
                    Cursor = Cursors.WaitCursor;
                    var result = new Dictionary<string, object?>();
                    var sqlAdHoc = _txtSqlAdHoc.Text;

                   
                    var resultQuery = await Task.Run(() => _reportService.ExecuteSQLAdHocAsync(sqlAdHoc, result, GetSelectedConnectionsAdHoc(), cts.Token));

                    if (resultQuery != null && resultQuery.Data != null)
                        _gridAdHoc.DataSource = resultQuery.Data;
                    PaintControlsTopGrid(_gridAdHoc, null, ResultTabUI.TabSecundary);

                   
                }
                catch (OperationCanceledException ex)
                {
                    // Cancelado = NO es error
                    _gridAdHoc.DataSource = null;
                    MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (OracleException ex) when (ex.Number == 942)
                {
                    GlobalExceptionHandler.Handle(ex, null,
                        "La tabla o vista no existe en la base de datos. Verifique que está ejecutando " +
                        "la sentencia correcta en la base de datos seleccionada.\n\n");
                    _gridAdHoc.DataSource = null;
                }
                catch (OracleException ex) when (ex.Number == 1013)
                {
                    // ORA-01013 = cancelación del usuario
                    _gridAdHoc.DataSource = null;
                    MessageBox.Show("Consulta cancelada por el usuario.",
                        "Cancelado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                finally
                {
                     RecursiveEnableControlsForm(this, true);
                    Cursor = Cursors.Default;
                    if (!loadingFormAdHoc.IsDisposed)
                        loadingFormAdHoc.Close();
                     Enabled = true;
                    _btnUnselectAllAdHoc.PerformClick();

                }
            }
        }

        private void RecursiveEnableControlsForm(Control control, bool changeStated)
        {
            if (control == null)
                return;
            if(control.Name.Contains("LoadingForm"))
                return;

            control.Enabled = changeStated;

            foreach (Control child in control.Controls)
                RecursiveEnableControlsForm(child, changeStated);
        }

        private async void MainForm_LoadAsync(object? sender, EventArgs e)
        {
            await LoadReportsAsync();
        }

        #endregion

        #region Inicialización de diseño

        private void ConfigurarTopPanel()
        {
            _topPanel.Controls.Add(_btnSelectAll);
            _topPanel.Controls.Add(_btnUnselectAll);
            // _topPanel.Controls.Add(_btnExport);

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
            _topPanel.Controls.Add(_btnVerConsulta);
            _btnSelectAll.Tag = ResultTabUI.TabInitial;
            _btnSelectAll.Click += BtnSelectAll_Click;
            _btnUnselectAll.Click += BtnUnselectAll_Click;
            _btnRunReport.Click += BtnRunReport_Click;
            _btnVerConsulta.Click += BtnVerConsulta_Click;
            _cmbReports.SelectedIndexChanged += CmbReports_SelectedIndexChanged;
        }

        private void ConfigurarGrupoParametros()
        {
            _grpParametros.Controls.Clear();
            _grpParametros.Controls.Add(_paramsPanel);
            _grpParametros.Height = 160;
        }

        #endregion

        #region Carga de datos

        private void CargarConexiones()
        {
            var conexiones = _connectionCatalog
                .GetAllConnections()
                .OrderBy(c => c.Type)
                .ThenBy(c => c.Id);

            _chkConnections.Items.Clear();

            var connectionCentral = conexiones.Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                x.DisplayName.IndexOf("Central", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            foreach (ConnectionInfo c in connectionCentral)
            {
                c.DisplayName = c.DisplayName!.ToUpperInvariant().Trim();
                _chkConnections.Items.Add(c, false);
            }

            var connectionStation = conexiones.Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                x.DisplayName.IndexOf("I.T.V.", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            foreach (ConnectionInfo c in connectionStation)
                _chkConnections.Items.Add(c, false);

            var connectionUma = conexiones.Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                x.DisplayName.IndexOf("U.M.A.", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            foreach (ConnectionInfo c in connectionUma)
                _chkConnections.Items.Add(c, false);

            _chkConnections.AutoAdjustWidth();
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

        #region Eventos de UI

        private void BtnSelectAll_Click(object? sender, EventArgs e)
        {
            if (sender != null && sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(true, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(true, _chkConnectionsAdHoc);
                        break;
                    default:
                        break;
                }
            }
        }

        private void BtnUnselectAll_Click(object? sender, EventArgs e)
        {
            if (sender != null && sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(false, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(false, _chkConnectionsAdHoc);
                        break;
                    default:
                        break;
                }
            }
        }

        private void CmbReports_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_cmbReports.SelectedItem is not ReportDefinition report)
                return;

            CargarConexiones();
            _grid.DataSource = null;

            if (_btnExcel != null)
                _btnExcel.Visible = false;

            if (_lblCountRows != null)
                _lblCountRows.Text = string.Empty;

            _currentReport = report;
            RenderParameters(report);
        }

        private async void BtnRunReport_Click(object? sender, EventArgs e)
        {
            var listConnectionsActive = GetSelectedConnections();

            if (listConnectionsActive == null || listConnectionsActive.Count == 0)
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

            if (_currentReport.Parameters != null &&
                _currentReport.Parameters.Count > 0 &&
                parametros.Count == 0)
            {
                return;
            }
            using var cts = new CancellationTokenSource();
            using var loading = new LoadingForm("Cargando datos...", cts);

            try
            {
                 RecursiveEnableControlsForm(this, false);
                loading.Owner = this;
                loading.Show();
                loading.Refresh();
                Enabled = false;
                Cursor = Cursors.WaitCursor;

                var resultReport = await Task.Run(() => _reportService.ExecuteReportAsync(
                    report,
                    parametros,
                    listConnectionsActive,
                    cts.Token));
                _grid.DataSource = resultReport.Data;
                PaintControlsTopGrid(_grid, resultReport, ResultTabUI.TabInitial);
                if (resultReport.TimeoutConnections.Any())
                {
                    var estaciones = string.Join(", ", resultReport.TimeoutConnections);
                    MessageBox.Show(
                        $"No se ha podido obtener información de las siguientes conexiones (timeout):{Environment.NewLine}{estaciones}",
                        "Aviso de Oracle",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (OperationCanceledException ex)
            {
                // Cancelado = NO es error
                _gridAdHoc.DataSource = null;
                MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OracleException ex) when (ex.Number == 1013)
            {
                // ORA-01013 = cancelación del usuario
                _gridAdHoc.DataSource = null;
                MessageBox.Show("Consulta cancelada por el usuario.",
                    "Cancelado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OracleException ex) when (ex.Number == 942)
            {
                GlobalExceptionHandler.Handle(ex, null,
                    "La tabla o vista no existe en la base de datos. Verifique que está ejecutando " +
                    "el informe correcto en la base de datos seleccionada.\n\n" +
                    "Revise la consulta mediante el botón \"Ver consulta\" \n" +
                    "Está ejecutando " +
                    $"un informe de {_currentReport?.SourceType}.");
                _grid.DataSource = null;
            }
            finally
            {
                RecursiveEnableControlsForm(this, true);
                Cursor = Cursors.Default;
                if (!loading.IsDisposed)
                    loading.Close();
                Enabled = true;
                _btnUnselectAll.PerformClick();
            }
        }



        private void RemoveControlsTopGrid(DataGridView grid,  ResultTabUI nameTab)
        {
            if (grid == null || grid.Parent == null)
                return;
            Control parent = grid.Parent;

            if (parent is TableLayoutPanel && parent.Parent != null)
                parent = parent.Parent;

            System.Drawing.Point gridLocationInParent = parent.PointToClient(
                grid.Parent.PointToScreen(grid.Location));

            var lblCountRowsExist = parent.Controls
                .OfType<Label>()
                .FirstOrDefault(l => l.Name == $"lblCountRows_{nameTab}");

            if (lblCountRowsExist != null)
                parent.Controls.Remove(lblCountRowsExist);

            var btnExcelExist = parent.Controls
                .OfType<Button>()
                .FirstOrDefault(b => b.Name == $"btnExportExcel_{nameTab}");
            if (btnExcelExist != null)
                parent.Controls.Remove(btnExcelExist);

                 grid.DataSource = null;
        }

        private void PaintControlsTopGrid(DataGridView grid, ReportQueryResult? resultReport, ResultTabUI nameTab)
        {
            if (grid == null || grid.Parent == null)
                return;
            // Padre  colocaremos label y botón
            Control parent = grid.Parent;

            // En la pestaña AdHoc el padre es un TableLayoutPanel, usamos su contenedor
            if (parent is TableLayoutPanel && parent.Parent != null)
                parent = parent.Parent;

            // Posición del grid relativa al padre
            System.Drawing.Point gridLocationInParent = parent.PointToClient(
                grid.Parent.PointToScreen(grid.Location));

            var lblCountRowsExist = parent.Controls
                .OfType<Label>()
                .FirstOrDefault(l => l.Name == $"lblCountRows_{nameTab}");

            if (lblCountRowsExist == null)
            {
                _lblCountRows = new Label
                {
                    Name = $"lblCountRows_{nameTab}",
                    AutoSize = true,
                    ForeColor = SystemColors.GrayText,
                    BackColor = Color.Transparent,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    Margin = new Padding(4)
                };
                parent.Controls.Add(_lblCountRows);
            }
            else
                _lblCountRows = lblCountRowsExist;

            int rowCount;
            if (resultReport != null && resultReport.Data != null)
                rowCount = resultReport.Data.Rows.Count;
            else
                rowCount = (grid.DataSource as DataTable)?.Rows.Count ?? 0;

            _lblCountRows.Text = $"Registros encontrados: {rowCount}";
            // el label encima del grid
            _lblCountRows.Top = gridLocationInParent.Y - _lblCountRows.Height - 15;
            _lblCountRows.Left = parent.ClientSize.Width - _lblCountRows.Width - 10;
            _lblCountRows.BringToFront();

            var btnExcelExist = parent.Controls
                .OfType<Button>()
                .FirstOrDefault(b => b.Name == $"btnExportExcel_{nameTab}");
            if (btnExcelExist == null)
            {
                _btnExcel = new Button
                {
                    Name = $"btnExportExcel_{nameTab}",
                    Size = new Size(24, 24),
                    FlatStyle = FlatStyle.Flat,
                    TabStop = false,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    Visible = true,
                    Tag = nameTab
                };

                _btnExcel.FlatAppearance.BorderSize = 0;
                var resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
                var iconObj = resources.GetObject("Excel_24");
                if (iconObj is Icon excelIcon)
                    _btnExcel.Image = excelIcon.ToBitmap();
                _btnExcel.Click += ExportGridWithClosedXml;
                parent.Controls.Add(_btnExcel);
            }
            else
                _btnExcel = btnExcelExist;

            _btnExcel.Top = _lblCountRows.Top - 3 + (_lblCountRows.Height - _btnExcel.Height) / 2;
            _btnExcel.Left = _lblCountRows.Left - _btnExcel.Width - 6;
            _btnExcel.Visible = rowCount > 0;
            _lblCountRows.BringToFront();
            _btnExcel.BringToFront();
        }
        private void BtnVerConsulta_Click(object? sender, EventArgs e)
        {
            if (_currentReport == null)
            {
                MessageBox.Show("No hay informe seleccionado.", "Ver consulta",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var initialConnection = _connectionCatalog
                .GetAllConnections()
                .FirstOrDefault(x => x.Type.ToUpper().Contains(_currentReport.SourceType.ToString().ToUpper()));

            if (initialConnection == null)
            {
                MessageBox.Show("No se encontró una conexión adecuada para ver la consulta.",
                    "Ver consulta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var con = _connectionFactory.CreateConnection(
                string.Concat(initialConnection.Id, "_", initialConnection.DisplayName));

            using (var frm = new SqlPreviewForm(_currentReport, con))
            {
                frm.ShowDialog(this);
            }
        }

        #endregion

        #region Métodos privados de ayuda

        private void SetAllConnectionsChecked(bool isChecked, CheckedListBox _chkCon)
        {
            for (int i = 0; i < _chkCon.Items.Count; i++)
            {
                _chkCon.SetItemChecked(i, isChecked);
            }
        }

        private void RenderParameters(ReportDefinition report)
        {
            _parameterControls.Clear();
            _paramsPanel.SuspendLayout();
            _paramsPanel.Controls.Clear();

            bool hasMaster = report.TableMasterForParameters != null && report.TableMasterForParameters.Count > 0;
            bool hasParams = report.Parameters != null && report.Parameters.Count > 0;

            if (!hasMaster && !hasParams)
            {
                var lbl = new Label
                {
                    Text = "Este informe no requiere parámetros.",
                    AutoSize = true,
                    Margin = new Padding(4, 8, 4, 4)
                };

                _paramsPanel.Controls.Add(lbl);
                _grpParametros.Height = 100;
                _paramsPanel.ResumeLayout();
                return;
            }

            FlowLayoutPanel CreateParamBlock(string labelText, Control input, bool? filterLike)
            {
                var block = new FlowLayoutPanel
                {
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = false,
                    Margin = new Padding(8, 4, 8, 4),
                    Padding = new Padding(0)
                };

                var lbl = new Label
                {
                    Text = labelText,
                    AutoSize = true,
                    Margin = new Padding(0, 6, 8, 0)
                };

                CheckBox? chkLike = null;
                if (filterLike != null && filterLike.Value)
                {
                    chkLike = new CheckBox
                    {
                        Checked = filterLike.Value,
                        AutoSize = true,
                        Name = "chkBusquedaLike",
                        Margin = new Padding(4, 4, 4, 2),
                        Text = "Búsqueda 'LIKE'"
                    };
                }

                input.Margin = new Padding(0, 2, 0, 0);
                if (input is CheckBox)
                    input.Margin = new Padding(0, 6, 0, 0);

                block.Controls.Add(lbl);
                block.Controls.Add(input);

                if (chkLike != null)
                    block.Controls.Add(chkLike);

                return block;
            }

            FlowLayoutPanel? lastMasterBlock = null;
            int masterCount = 0;

            if (hasMaster)
            {
                foreach (var p in report.TableMasterForParameters!)
                {
                    Control? input = CreateControlForTableMasterParameter(p, report.SourceType);
                    if (input == null) continue;

                    if (input is ListBox or CheckedListBox)
                    {
                        input.Width = 420;
                        input.Height = 140;
                    }

                    var block = CreateParamBlock(p.Label ?? p.Name, input, false);
                    _paramsPanel.Controls.Add(block);
                    masterCount++;
                    lastMasterBlock = block;
                    _parameterControls[p.Name] = input;

                    if (masterCount % 2 == 0)
                    {
                        input.Margin = new Padding(
                            input.Margin.Left,
                            input.Margin.Top,
                            input.Margin.Right,
                            10);
                        _paramsPanel.SetFlowBreak(block, true);
                    }
                }

                if (lastMasterBlock != null)
                {
                    lastMasterBlock.Margin = new Padding(
                        lastMasterBlock.Margin.Left,
                        lastMasterBlock.Margin.Top,
                        lastMasterBlock.Margin.Right,
                        10);
                    _paramsPanel.SetFlowBreak(lastMasterBlock, true);
                }
            }

            if (hasParams)
            {
                foreach (var p in report.Parameters!)
                {
                    Control? input = CreateControlForParameter(p);
                    if (input == null) continue;

                    if (input is DateTimePicker)
                        input.Width = 110;
                    else if (input is TextBox)
                        input.Width = 160;

                    var block = CreateParamBlock(p.Label ?? p.Name, input, p.BusquedaLike);
                    _paramsPanel.Controls.Add(block);
                    _parameterControls[p.Name] = input;
                }
            }

            _paramsPanel.ResumeLayout();

            if (_paramsPanel.Controls.Count > 0)
            {
                int maxBottom = _paramsPanel.Controls.Cast<Control>().Max(c => c.Bottom);
                _grpParametros.Height = maxBottom + 60;
            }
            else
            {
                _grpParametros.Height = 100;
            }
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
                case "text":
                    return new TextBox
                    {
                        Text = parameter.IsRequired ? string.Empty : "",
                        Width = 160,
                        Margin = new Padding(4, 4, 4, 2)
                    };
                case "bool":
                case "funcion":
                    return new CheckBox
                    {
                        Checked = parameter.IsRequired,
                        AutoSize = true,
                        Margin = new Padding(4, 4, 4, 2)
                    };
                default:
                    return null;
            }
        }

        private Control? CreateControlForTableMasterParameter(TableMasterParameterDefinition parameter, ReportSourceType sourceType)
        {
            var type = (parameter.Type ?? "string").ToLowerInvariant();

            return type switch
            {
                "combobox" => LoadTableMasterDataIntoControl(parameter, sourceType),
                _ => null
            };
        }

        private Control LoadTableMasterDataIntoControl(TableMasterParameterDefinition parameter, ReportSourceType sourceType)
        {
            var initialConnection = _connectionCatalog
                .GetAllConnections()
                .FirstOrDefault(x => x.Type.Contains(sourceType.ToString()));

            if (initialConnection == null)
                throw new Exception("No se encontró una conexión válida para Estación.");

            if (string.IsNullOrWhiteSpace(parameter.SqlQueryMaster))
                return new CheckedListBox();

            DataTable dt;

            using (var conn = _connectionFactory.CreateConnection(
                string.Concat(initialConnection.Id, "_", initialConnection.DisplayName)) as OracleConnection)
            {
                conn!.Open();
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
                Height = Math.Min(150, 18 * dt.Rows.Count + 4)
            };

            clb.DisplayMember = parameter.Text ?? string.Empty;
            clb.ValueMember = parameter.Id ?? string.Empty;

            var preselected = parameter.ValuesRequired ?? new List<string>();

            int maxWidth = 0;

            foreach (DataRow row in dt.Rows)
            {
                string value = row[parameter.Id]?.ToString() ?? "";
                string text = row[parameter.Text]?.ToString() ?? "";

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

        private List<ConnectionInfo> GetSelectedConnections()
        {
            return _chkConnections.CheckedItems
                .OfType<ConnectionInfo>()
                .ToList();
        }

        private List<ConnectionInfo> GetSelectedConnectionsAdHoc()
        {
            return _chkConnectionsAdHoc.CheckedItems
                .OfType<ConnectionInfo>()
                .ToList();
        }

        private Dictionary<string, object?> BuildParametersFromUI()
        {
            var result = new Dictionary<string, object?>();

            if (_currentReport == null)
                return result;

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

                    if (value != null && value is bool)
                    {
                        if (p.Type == "funcion")
                        {
                            int numberFromBoolean = value is bool b ? (b ? 1 : 0) : 0;
                            value = p.Values
                                .Where(x => x.Key == numberFromBoolean)
                                .FirstOrDefault()
                                .Value ?? string.Empty;
                        }
                        else
                        {
                            value = (bool)value ? -1 : 0;
                        }
                    }

                    if (p.Type == "text")
                    {
                        if (ctrl.Parent is FlowLayoutPanel flp)
                        {
                            var chkLike = flp.Controls.OfType<CheckBox>()
                                .FirstOrDefault(c => c.Name == "chkBusquedaLike");
                            if (chkLike != null)
                            {
                                if (value != null)
                                {
                                    value = string.Concat("%", value.ToString()!.Trim(), "%");
                                    ReplaceSqlInput(p.Name, _currentReport, value);
                                }
                                else
                                    value = "%%";
                            }
                            else
                            {
                                if (value != null)
                                {
                                    value = value.ToString()!.Trim();
                                    ReplaceSqlInput(p.Name, _currentReport, value);
                                }
                                else
                                    value = "%%";
                            }
                        }
                        else
                            value = value?.ToString()!.Trim();
                    }

                    switch (p.Name.ToUpper())
                    {
                        case "FECHADESDE":
                            var fromDate = (DateTime)value!;
                            value = new DateTime(fromDate.Year, fromDate.Month, fromDate.Day);
                            break;

                        case "FECHAHASTA":
                            var toDate = (DateTime)value!;
                            value = new DateTime(toDate.Year, toDate.Month, toDate.Day, 23, 59, 59);
                            break;
                    }

                    result[p.Name] = value;
                }
            }

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
                    return "''";

                return string.Join(", ",
                    values.Select(v => $"'{v.Replace("'", "''")}'"));
            }

            var tiposVehiculo = GetCheckedCodes("TIPOSVEHICULO");
            if (tiposVehiculo != null && tiposVehiculo.Count > 0)
                result["TiposVehiculoList"] = BuildInList(tiposVehiculo);

            var categorias = GetCheckedCodes("CATEGORIAS");
            if (categorias != null && categorias.Count > 0)
                result["CategoriasList"] = BuildInList(categorias);

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

        private void ReplaceSqlInput(string nameParameter, ReportDefinition? currentReport, object? value)
        {
            string pattern = ":" + nameParameter + @"([^)]*)\)";
            string patternIsnull = ":" + nameParameter + @"([^)]*)\)";

            if (currentReport?.SourceType == ReportSourceType.Central && !string.IsNullOrEmpty(currentReport.SqlForCentral))
            {
                if (value != null && value?.ToString() == "%%")
                {
                    var matchIsnull = Regex.Match(currentReport.SqlForCentral, patternIsnull);
                    var txtReplace = matchIsnull.Value.Replace($"{nameParameter}", string.Empty);
                    currentReport.SqlForCentral = Regex.Replace(currentReport.SqlForCentral, pattern, $":{txtReplace} )");
                }
                else
                    currentReport.SqlForCentral = Regex.Replace(currentReport.SqlForCentral, pattern, $":{nameParameter})");
            }
            else if (currentReport?.SourceType == ReportSourceType.Estacion && !string.IsNullOrEmpty(currentReport.SqlForStations))
            {
                if (value != null && value?.ToString() == "%%")
                {
                    var matchIsnull = Regex.Match(currentReport.SqlForStations, patternIsnull);
                    var txtReplace = matchIsnull.Value.Replace($"{nameParameter}", string.Empty);
                    currentReport.SqlForCentral = Regex.Replace(currentReport.SqlForStations, pattern, $":{txtReplace} )");
                }
                else
                    currentReport.SqlForStations = Regex.Replace(currentReport.SqlForStations, pattern, $":{nameParameter})");
            }
        }

        private void ExportGridWithClosedXml(object? sender, EventArgs e)
        {
            using var cts = new CancellationTokenSource();
            using var loading = new LoadingForm("Exportando datos a Excel ...", cts);
            try
            {
                RecursiveEnableControlsForm(this, false);
                loading.Owner = this;
                loading.Show();
                loading.Refresh();
                var uniqueIdTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                using var sfd = new SaveFileDialog
                {
                    Filter = "Excel (*.xlsx)|*.xlsx"
                };
                if (sender != null && sender is Button btn && btn.Tag is ResultTabUI typeExport)
                {
                    switch (typeExport)
                    {
                        case ResultTabUI.TabInitial:
                            sfd.FileName = $"{((ReportDefinition)_cmbReports.SelectedItem!).Name}_{uniqueIdTime}.xlsx";
                            break;
                        case ResultTabUI.TabSecundary:
                            sfd.FileName = $"ConsultaPersonalizada_{uniqueIdTime}.xlsx";
                            break;
                        default:
                            break;
                    }
                    if (sfd.ShowDialog() != DialogResult.OK)
                        return;
                    IXLWorksheet? ws = null;
                    using var wb = new XLWorkbook();
                    switch (typeExport)
                    {
                        case ResultTabUI.TabInitial:
                            ws = wb.Worksheets.Add(((ReportDefinition)_cmbReports.SelectedItem!).Category);
                            break;
                        case ResultTabUI.TabSecundary:
                            ws = wb.Worksheets.Add("ConsultaPersonalizada");
                            break;
                    }
                    if (ws == null)
                        throw new Exception("Error al crear el Excel");
                    DataGridView? _gridExport = null;
                    if (typeExport == ResultTabUI.TabInitial)
                        _gridExport = _grid;
                    else if (typeExport == ResultTabUI.TabSecundary)
                        _gridExport = _gridAdHoc;

                    for (int col = 0; col < _gridExport?.Columns.Count; col++)
                    {
                        ws.Cell(1, col + 1).Value = _gridExport.Columns[col].HeaderText;
                        ws.Cell(1, col + 1).Style.Font.Bold = true;
                    }
                    for (int row = 0; row < _gridExport?.Rows.Count; row++)
                    {
                        if (_gridExport.Rows[row].IsNewRow) continue;
                        for (int col = 0; col < _gridExport.Columns.Count; col++)
                        {
                            var value = _gridExport.Rows[row].Cells[col].Value;
                            string ? safeValue = value == null ? "" : value.ToString();
                            ws.Cell(row + 2, col + 1).Value = safeValue?.Trim();
                        }
                    }
                    ws.Columns().AdjustToContents();
                    foreach (var sheet in wb.Worksheets)
                    {
                        sheet.Columns().AdjustToContents();
                    }
                    wb.SaveAs(sfd.FileName);
                    MessageBox.Show(
                        $"Exportación realizada correctamente en:\n{sfd.FileName}",
                        "Exportación informe",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show(
                    "El archivo puede estar abierto. Por favor, ciérrelo para poder guardar el archivo Excel.",
                    "Error Exportación Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                RecursiveEnableControlsForm(this, true);
                Cursor = Cursors.Default;
                if(!loading.IsDisposed)
                        loading.Close();
                Enabled = true;
            }
        }

        #endregion

        #region Form de carga

        private sealed class LoadingForm : Form
        {
            private readonly CancellationTokenSource? _cts;
            private  Panel _buttonsPanel;
            private Button btnCancelar;

            public LoadingForm(string message, CancellationTokenSource? cts = null)
            {
                _cts = cts;
                Name = "LoadingForm";
                StartPosition = FormStartPosition.Manual;
                TopMost = true;
                ShowInTaskbar = false;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                ControlBox = false;
                Width = 260;
                Height = 120; 
                Text = string.Empty;

                var lbl = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Text = message,
                    Font = new Font(SystemFonts.DefaultFont.FontFamily, 10, FontStyle.Bold)
                };

                Controls.Add(lbl);

                _buttonsPanel = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = 42
                };
                Controls.Add(_buttonsPanel);

                // Si nos pasan un CTS, añadimos botón Cancelar
                if (_cts != null)
                {
                      btnCancelar = new Button
                    {
                        Text = "Cancelar",
                        AutoSize = true,
                        Height = 28,
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font(SystemFonts.DefaultFont.FontFamily, 9f, FontStyle.Bold),
                        BackColor = Color.FromArgb(230, 230, 230),
                        ForeColor = Color.FromArgb(50, 50, 50),
                        Cursor = Cursors.Hand,
                        TabStop = false,
                        Padding = new Padding(12, 4, 12, 4)
                    };

                    btnCancelar.FlatAppearance.BorderSize = 1;
                    btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
                    btnCancelar.Click += (_, __) => _cts.Cancel();

                    //Controls.Add(btnCancelar);
                    _buttonsPanel.Controls.Add(btnCancelar);
                    CenterButton();

                    _buttonsPanel.Resize += (_, __) => CenterButton();
                }
            }
            private void CenterButton()
            {
                if (btnCancelar == null) return;

                btnCancelar.Left = (_buttonsPanel.Width - btnCancelar.Width) / 2;
                btnCancelar.Top = (_buttonsPanel.Height - btnCancelar.Height) / 2;
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);

                if (Owner != null)
                {
                    var rect = Owner.Bounds;
                    Left = rect.Left + (rect.Width - Width) / 2;
                    Top = rect.Top + (rect.Height - Height) / 2;
                }
                else
                {
                    var screen = Screen.FromPoint(Cursor.Position).WorkingArea;
                    Left = screen.Left + (screen.Width - Width) / 2;
                    Top = screen.Top + (screen.Height - Height) / 2;
                }

                CenterButton();

            }


        }

        #endregion

    }

    #region Clases auxiliares

    public class MultiItem
    {
        public string Value { get; set; } = "";
        public string Text { get; set; } = "";

        public override string ToString()
            => $"({Value}) -> {Text}";
    }

    #endregion
}
