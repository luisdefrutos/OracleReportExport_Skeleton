using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using Spre = DocumentFormat.OpenXml.Spreadsheet;
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
using System.Threading;
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
        private static readonly Regex RegexParams = new(@"(?<!:):(?!\d)\w+", RegexOptions.Compiled);

        // Paginadores lógicos (uno por pestaña)
        private PropertyGrid? _pagerPredef;
        private PropertyGrid? _pagerAdHoc;

        // Botones de paginación AdHoc
        private readonly Button _btnPrevPageAdHoc = new()
        {
            Name = "_btnPrevPageAdHoc",
            Text = "< Anterior",
            AutoSize = true,
            Visible = false,
            Enabled = false
        };

        private readonly Button _btnNextPageAdHoc = new()
        {
            Name = "_btnNextPageAdHoc",
            Text = "Siguiente >",
            AutoSize = true,
            Visible = false,
            Enabled = false
        };

        // Botones de paginación Predefinidos
        private readonly Button _btnPrevPagePredef = new()
        {
            Name = "_btnPrevPagePredef",
            Text = "< Anterior",
            AutoSize = true,
            Enabled = false
        };

        private readonly Button _btnNextPagePredef = new()
        {
            Name = "_btnNextPagePredef",
            Text = "Siguiente >",
            AutoSize = true,
            Enabled = false
        };

        // Labels de estado por pestaña
        private readonly Label _lblCountRowsPredef = new()
        {
            AutoSize = true,
            ForeColor = SystemColors.ControlText,   // misma letra que el resto de la app
            Margin = new Padding(4),
            Visible = false
        };

        private readonly Label _lblCountRowsAdHoc = new()
        {
            AutoSize = true,
            ForeColor = SystemColors.ControlText,
            Margin = new Padding(4),
            Visible = false
        };

        // Botones Excel por pestaña
        private readonly Button _btnExcelPredef = new()
        {
            Name = "btnExportExcel_Predef",
            Size = new Size(24, 24),
            FlatStyle = FlatStyle.Flat,
            TabStop = false,
            Visible = false,
            Tag = ResultTabUI.TabInitial
        };

        private readonly Button _btnExcelAdHoc = new()
        {
            Name = "btnExportExcel_AdHoc",
            Size = new Size(24, 24),
            FlatStyle = FlatStyle.Flat,
            TabStop = false,
            Visible = false,
            Tag = ResultTabUI.TabSecundary
        };

        // Grids principales
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

        // Conexiones
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

        // Botones de selección de conexiones
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

        // Paneles superiores
        private readonly FlowLayoutPanel _topPanel = new()
        {
            Dock = DockStyle.Top,
            Height = 52,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(5, 5, 5, 5),
            Margin = new Padding(0, 0, 0, 10)
        };

        private readonly FlowLayoutPanel _topPanelAdHoc = new()
        {
            Dock = DockStyle.Top,
            Height = 52,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(5, 5, 5, 5),
            Margin = new Padding(0, 0, 0, 10)
        };

        // Parámetros
        private readonly GroupBox _grpParametros = new()
        {
            Text = "Parámetros",
            Dock = DockStyle.Top,
            Height = 130,
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

        // SQL AdHoc
        private readonly RichTextBox _txtSqlAdHoc = new()
        {
            Dock = DockStyle.Fill,
            Font = new System.Drawing.Font("Consolas", 10f),
            BackColor = System.Drawing.Color.FromArgb(232, 238, 247),
            ForeColor = System.Drawing.Color.FromArgb(28, 59, 106),
            ScrollBars = RichTextBoxScrollBars.Both,
            BorderStyle = BorderStyle.FixedSingle,
            ShortcutsEnabled = true
        };

        private readonly Button ButtonAdHoc = new()
        {
            AutoSize = true,
            Text = "Ejecutar SQL",
            Visible = true
        };

        private readonly Dictionary<string, Control> _parameterControls = new();

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

            var resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));

            // ------------------------------------------------------------------
            // PESTAÑA INFORMES PREDEFINIDOS
            // ------------------------------------------------------------------

            var rightPredefLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4
            };
            rightPredefLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // parámetros
            rightPredefLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // info (icono+texto)
            rightPredefLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));   // grid
            rightPredefLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // paginación

            // 0) Parámetros
            rightPredefLayout.Controls.Add(_grpParametros, 0, 0);

            // 1) Panel superior info (icono + texto) alineado a la derecha
            var infoPredefOuter = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 26
            };
            var infoPredefRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(5, 3, 5, 0),
                Margin = new Padding(0)
            };

            _btnExcelPredef.FlatAppearance.BorderSize = 0;
            var iconObj1 = resources.GetObject("Excel_24");
            if (iconObj1 is Icon excelIcon1)
                _btnExcelPredef.Image = excelIcon1.ToBitmap();

            infoPredefRight.Controls.Add(_btnExcelPredef);     // primero icono
            infoPredefRight.Controls.Add(_lblCountRowsPredef); // luego texto

            infoPredefOuter.Controls.Add(infoPredefRight);
            rightPredefLayout.Controls.Add(infoPredefOuter, 0, 1);

            // 2) Grid
            _grid.Dock = DockStyle.Fill;
            rightPredefLayout.Controls.Add(_grid, 0, 2);

            // 3) Panel inferior de paginación
            var pagerPredefOuter = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 32
            };
            var pagerPredefRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(5, 3, 5, 3),
                Margin = new Padding(0)
            };

            pagerPredefRight.Controls.Add(_btnPrevPagePredef);
            pagerPredefRight.Controls.Add(_btnNextPagePredef);

            pagerPredefOuter.Controls.Add(pagerPredefRight);
            rightPredefLayout.Controls.Add(pagerPredefOuter, 0, 3);

            var mainPredefPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            _chkConnections.Dock = DockStyle.Left;
            _chkConnections.Width = 260;

            mainPredefPanel.Controls.Add(rightPredefLayout);
            mainPredefPanel.Controls.Add(_chkConnections);

            _tabPredefinidos.Controls.Add(mainPredefPanel);
            _tabPredefinidos.Controls.Add(_topPanel);

            // ------------------------------------------------------------------
            // PESTAÑA SQL AVANZADA (ADHOC)
            // ------------------------------------------------------------------

            _chkConnectionsAdHoc.Items.AddRange(
                _chkConnections.Items.OfType<ConnectionInfo>().ToArray()
            );

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
            _topPanelAdHoc.Controls.Add(ButtonAdHoc); // botón de ejecución el más a la derecha

            var rightPanelAdHoc = new Panel
            {
                Dock = DockStyle.Fill
            };

            var layoutAdHoc = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4
            };
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));   // SQL
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // info (icono+texto)
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 60f));   // grid
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // paginación

            // 0) editor SQL
            _txtSqlAdHoc.Dock = DockStyle.Fill;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Copiar", null, TxtSql_Copy_Click);
            menu.Items.Add("Pegar", null, TxtSql_Paste_Click);
            menu.Items.Add("Cortar", null, TxtSql_Cut_Click);
            menu.Items.Add("Seleccionar todo", null, TxtSql_SelectAll_Click);
            _txtSqlAdHoc.ContextMenuStrip = menu;

            layoutAdHoc.Controls.Add(_txtSqlAdHoc, 0, 0);

            // 1) panel info AdHoc (icono + texto) alineado a la derecha
            var infoAdHocOuter = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 26
            };
            var infoAdHocRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(5, 3, 5, 0),
                Margin = new Padding(0)
            };

            _btnExcelAdHoc.FlatAppearance.BorderSize = 0;
            var iconObj2 = resources.GetObject("Excel_24");
            if (iconObj2 is Icon excelIcon2)
                _btnExcelAdHoc.Image = excelIcon2.ToBitmap();

            infoAdHocRight.Controls.Add(_btnExcelAdHoc);      // icono
            infoAdHocRight.Controls.Add(_lblCountRowsAdHoc);  // texto
            infoAdHocOuter.Controls.Add(infoAdHocRight);
            layoutAdHoc.Controls.Add(infoAdHocOuter, 0, 1);

            // 2) grid resultados
            _gridAdHoc.Dock = DockStyle.Fill;
            layoutAdHoc.Controls.Add(_gridAdHoc, 0, 2);

            // 3) panel inferior de paginación AdHoc
            var pagerAdHocOuter = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 32
            };
            var pagerAdHocRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(5, 3, 5, 3),
                Margin = new Padding(0)
            };

            pagerAdHocRight.Controls.Add(_btnPrevPageAdHoc);
            pagerAdHocRight.Controls.Add(_btnNextPageAdHoc);

            pagerAdHocOuter.Controls.Add(pagerAdHocRight);
            layoutAdHoc.Controls.Add(pagerAdHocOuter, 0, 3);

            rightPanelAdHoc.Controls.Add(layoutAdHoc);

            _chkConnectionsAdHoc.Dock = DockStyle.Left;
            _chkConnectionsAdHoc.Width = 260;

            tabAdHoc.Controls.Add(rightPanelAdHoc);
            tabAdHoc.Controls.Add(_chkConnectionsAdHoc);
            tabAdHoc.Controls.Add(_topPanelAdHoc);

            _tabControl.TabPages.Add(_tabPredefinidos);
            _tabControl.TabPages.Add(tabAdHoc);

            Controls.Add(_tabControl);

            // Eventos paginación
            _btnPrevPageAdHoc.Click += _btnPrevPageAdHoc_Click;
            _btnNextPageAdHoc.Click += _btnNextPageAdHoc_Click;
            _btnPrevPagePredef.Click += _btnPrevPagePredef_Click;
            _btnNextPagePredef.Click += _btnNextPagePredef_Click;

            // Excel
            _btnExcelPredef.Click += ExportGridWithClosedXml;
            _btnExcelAdHoc.Click += ExportGridWithClosedXml;

            Load += MainForm_LoadAsync;
        }

        #endregion

        #region Eventos UI básicos (context menu SQL)

        private void TxtSql_Copy_Click(object sender, EventArgs e) => _txtSqlAdHoc.Copy();
        private void TxtSql_Paste_Click(object sender, EventArgs e) => _txtSqlAdHoc.Paste();
        private void TxtSql_Cut_Click(object sender, EventArgs e) => _txtSqlAdHoc.Cut();
        private void TxtSql_SelectAll_Click(object sender, EventArgs e) => _txtSqlAdHoc.SelectAll();

        #endregion

        #region Eventos paginación

        private void _btnNextPageAdHoc_Click(object? sender, EventArgs e)
        {
            if (_pagerAdHoc == null) return;
            _pagerAdHoc.ShowNextPage();
            ActualizarPaginacionAdHoc();
        }

        private void _btnPrevPageAdHoc_Click(object? sender, EventArgs e)
        {
            if (_pagerAdHoc == null) return;
            _pagerAdHoc.ShowPreviousPage();
            ActualizarPaginacionAdHoc();
        }

        private void _btnNextPagePredef_Click(object? sender, EventArgs e)
        {
            if (_pagerPredef == null) return;
            _pagerPredef.ShowNextPage();
            ActualizarPaginacionPredefinidos();
        }

        private void _btnPrevPagePredef_Click(object? sender, EventArgs e)
        {
            if (_pagerPredef == null) return;
            _pagerPredef.ShowPreviousPage();
            ActualizarPaginacionPredefinidos();
        }

        #endregion

        #region Inicialización de diseño

        private void ConfigurarTopPanel()
        {
            _topPanel.Controls.Add(_btnSelectAll);
            _topPanel.Controls.Add(_btnUnselectAll);

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

            // Queremos ejecutar informe a la derecha del todo:
            _topPanel.Controls.Add(_btnVerConsulta);
            _topPanel.Controls.Add(_btnRunReport);

            _btnSelectAll.Tag = ResultTabUI.TabInitial;
            _btnSelectAll.Click += BtnSelectAll_Click;
            _btnUnselectAll.Tag = ResultTabUI.TabInitial;
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

            var connectionCentral = conexiones
                .Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                            x.DisplayName.IndexOf("Central", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            foreach (ConnectionInfo c in connectionCentral)
            {
                c.DisplayName = c.DisplayName!.ToUpperInvariant().Trim();
                _chkConnections.Items.Add(c, false);
            }

            var connectionStation = conexiones
                .Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                            x.DisplayName.IndexOf("I.T.V.", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            foreach (ConnectionInfo c in connectionStation)
                _chkConnections.Items.Add(c, false);

            var connectionUma = conexiones
                .Where(x => !string.IsNullOrEmpty(x.DisplayName) &&
                            x.DisplayName.IndexOf("U.M.A.", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

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

        private async void MainForm_LoadAsync(object? sender, EventArgs e)
        {
            await LoadReportsAsync();
        }

        #endregion

        #region Eventos de UI (conexiones y combo)

        private void BtnSelectAll_Click(object? sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(true, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(true, _chkConnectionsAdHoc);
                        break;
                }
            }
        }

        private void BtnUnselectAll_Click(object? sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(false, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(false, _chkConnectionsAdHoc);
                        break;
                }
            }
        }

        private void _btnSelectAllAdHoc_Click(object? sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(true, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(true, _chkConnectionsAdHoc);
                        break;
                }
            }
        }

        private void _btnUnselectAllAdHoc_Click(object? sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is ResultTabUI typeExport)
            {
                switch (typeExport)
                {
                    case ResultTabUI.TabInitial:
                        SetAllConnectionsChecked(false, _chkConnections);
                        break;
                    case ResultTabUI.TabSecundary:
                        SetAllConnectionsChecked(false, _chkConnectionsAdHoc);
                        break;
                }
            }
        }

        private void _btnClearAdHoc_Click(object? sender, EventArgs e)
        {
            _txtSqlAdHoc?.Clear();
        }

        private void CmbReports_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_cmbReports.SelectedItem is not ReportDefinition report)
                return;

            CargarConexiones();
            _grid.DataSource = null;

            _pagerPredef = null;
            ResetPaginacionPredefinidos();

            _currentReport = report;
            RenderParameters(report);
        }

        #endregion

        #region Ejecución informes predefinidos

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
                _pagerPredef = null;
                _grid.DataSource = null;
                ResetPaginacionPredefinidos();

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

                if (resultReport != null && resultReport.Data != null)
                {
                    _pagerPredef = new PropertyGrid(
                        resultReport.Data,
                        _grid,
                        ResultTabUI.TabInitial,
                        (ReportDefinition)_cmbReports.SelectedItem);

                    _pagerPredef.PageSize = 500;
                    _pagerPredef.PageChanged += PagerPredef_PageChanged;
                    _pagerPredef.ShowFirstPage();
                    ActualizarPaginacionPredefinidos();
                }

                if (resultReport != null && resultReport.TimeoutConnections.Any())
                {
                    var estaciones = string.Join(", ", resultReport.TimeoutConnections);
                    MessageBox.Show(
                        $"No se ha podido obtener información de las siguientes conexiones (timeout):{Environment.NewLine}{estaciones}",
                        "Aviso de Oracle",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (OperationCanceledException)
            {
                _grid.DataSource = null;
                MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OracleException ex) when (ex.Number == 1013)
            {
                _grid.DataSource = null;
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

        private void PagerPredef_PageChanged(object? sender, EventArgs e)
        {
            ActualizarPaginacionPredefinidos();
        }

        #endregion

        #region Ejecución SQL AdHoc

        private async Task<bool> ErrorSintaxAdHoc_Click(object? sender, EventArgs e)
        {
            bool stopProcess = false;

            var listConnectionsActive = GetSelectedConnectionsAdHoc();

            if (listConnectionsActive == null || listConnectionsActive.Count == 0)
            {
                MessageBox.Show(
                    "Selecciona al menos una conexión para ejecutar el informe.",
                    "Sin selección",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return true;
            }

            var sqlAdHocRichTxt = _txtSqlAdHoc.Text;
            if (string.IsNullOrWhiteSpace(sqlAdHocRichTxt))
            {
                MessageBox.Show(
                    "Introduce una sentencia SQL.",
                    "SQL vacía",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return true;
            }

            var paramNames = DetectarParametros(sqlAdHocRichTxt);

            using var cts = new CancellationTokenSource();
            using (var loadingFormAdHoc = new LoadingForm("Validando sentencia ....", cts))
            {
                try
                {
                    _pagerAdHoc = null;
                    _gridAdHoc.DataSource = null;
                    ResetPaginacionAdHoc();

                    RecursiveEnableControlsForm(this, false);
                    loadingFormAdHoc.Owner = this;
                    loadingFormAdHoc.Show();
                    loadingFormAdHoc.Refresh();
                    Enabled = false;
                    Cursor = Cursors.WaitCursor;

                    if (paramNames.Count == 0)
                    {
                        var connectionForValidation = listConnectionsActive.First();

                        stopProcess = await ValidarSqlSinParametrosAsync(sqlAdHocRichTxt, connectionForValidation, cts.Token);

                        if (stopProcess)
                            return true;
                        else
                        {
                            RecursiveEnableControlsForm(this, true);
                            Cursor = Cursors.Default;
                            if (!loadingFormAdHoc.IsDisposed)
                                loadingFormAdHoc.Close();
                            Enabled = true;

                            var resp = MessageBox.Show(
                                "La sentencia SQL es sintácticamente correcta.\n\n¿Desea ejecutarla?",
                                "Validación correcta",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            stopProcess = resp == DialogResult.No;
                        }
                    }

                    return stopProcess;
                }
                catch (OperationCanceledException)
                {
                    MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                catch (OracleException ex) when (ex.Number == 942)
                {
                    GlobalExceptionHandler.Handle(ex, null,
                        "La tabla o vista no existe en la sentencia ejecutada. Verifique que está ejecutando " +
                        "la sentencia correcta en la base de datos seleccionada.\n\n");
                    return true;
                }
                catch (OracleException ex) when (ex.Number == 1013)
                {
                    _gridAdHoc.DataSource = null;
                    MessageBox.Show("Consulta cancelada por el usuario.",
                        "Cancelado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return true;
                }
                finally
                {
                    RecursiveEnableControlsForm(this, true);
                    Cursor = Cursors.Default;
                    if (!loadingFormAdHoc.IsDisposed)
                        loadingFormAdHoc.Close();
                    Enabled = true;
                }
            }
        }

        private async void ButtonAdHoc_Click(object? sender, EventArgs e)
        {
            var resultSintax = await ErrorSintaxAdHoc_Click(sender, e);
            if (resultSintax)
                return;

            using var cts = new CancellationTokenSource();
            using (var loadingFormAdHoc = new LoadingForm("Cargando Datos Consulta ....", cts))
            {
                try
                {
                    _pagerAdHoc = null;
                    _gridAdHoc.DataSource = null;
                    ResetPaginacionAdHoc();

                    RecursiveEnableControlsForm(this, false);
                    loadingFormAdHoc.Owner = this;
                    loadingFormAdHoc.Show();
                    loadingFormAdHoc.Refresh();
                    Enabled = false;
                    Cursor = Cursors.WaitCursor;

                    var result = new Dictionary<string, object?>();
                    var sqlAdHoc = _txtSqlAdHoc.Text;

                    var resultQuery = await Task.Run(() =>
                        _reportService.ExecuteSQLAdHocAsync(sqlAdHoc, result, GetSelectedConnectionsAdHoc(), cts.Token));

                    if (resultQuery != null && resultQuery.Data != null)
                    {
                        _pagerAdHoc = new PropertyGrid(resultQuery.Data, _gridAdHoc, ResultTabUI.TabSecundary, null);
                        _pagerAdHoc.PageSize = 500;
                        _pagerAdHoc.PageChanged += PagerAdHoc_PageChanged;
                        _pagerAdHoc.ShowFirstPage();
                        ActualizarPaginacionAdHoc();
                    }
                }
                catch (OperationCanceledException)
                {
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
                        "la sentencia correcta en la base de datos seleccionada.\n\n");
                    _gridAdHoc.DataSource = null;
                }
                catch (OracleException ex) when (ex.Number == 1013)
                {
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

        private void PagerAdHoc_PageChanged(object? sender, EventArgs e)
        {
            ActualizarPaginacionAdHoc();
        }

        private async Task<bool> ValidarSqlSinParametrosAsync(string sql, ConnectionInfo connectionForValidation, CancellationToken ct)
        {
            try
            {
                var result = await _reportService.ValidateSqlSyntaxAsync(sql, connectionForValidation, ct);
                return result;
            }
            catch (OracleException)
            {
                throw;
            }
        }

        private List<string> DetectarParametros(string sql)
        {
            var matches = RegexParams.Matches(sql);

            return matches
                .Cast<Match>()
                .Select(m => m.Value)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        #endregion

        #region Métodos privados de ayuda (UI, parámetros, paginación)

        private void RecursiveEnableControlsForm(Control control, bool changeStated)
        {
            if (control == null)
                return;

            if (control.Name.Contains("LoadingForm"))
                return;

            if (control.Name.Contains("_btnPrevPageAdHoc") ||
                control.Name.Contains("_btnNextPageAdHoc") ||
                control.Name.Contains("_btnPrevPagePredef") ||
                control.Name.Contains("_btnNextPagePredef"))
                return;

            control.Enabled = changeStated;

            foreach (Control child in control.Controls)
                RecursiveEnableControlsForm(child, changeStated);
        }

        private void SetAllConnectionsChecked(bool isChecked, CheckedListBox chk)
        {
            for (int i = 0; i < chk.Items.Count; i++)
                chk.SetItemChecked(i, isChecked);
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

        private void ResetPaginacionPredefinidos()
        {
            _lblCountRowsPredef.Visible = false;
            _lblCountRowsPredef.Text = string.Empty;

            _btnExcelPredef.Visible = false;
            _btnPrevPagePredef.Enabled = false;
            _btnNextPagePredef.Enabled = false;
        }

        private void ResetPaginacionAdHoc()
        {
            _lblCountRowsAdHoc.Visible = false;
            _lblCountRowsAdHoc.Text = string.Empty;

            _btnExcelAdHoc.Visible = false;
            _btnPrevPageAdHoc.Enabled = false;
            _btnNextPageAdHoc.Enabled = false;
            _btnPrevPageAdHoc.Visible = false;
            _btnNextPageAdHoc.Visible = false;
        }

        private void ActualizarPaginacionPredefinidos()
        {
            if (_pagerPredef == null)
            {
                ResetPaginacionPredefinidos();
                return;
            }

            int totalFilas = _pagerPredef.TotalRows;
            var tablaPagina = _grid.DataSource as DataTable;
            int filasPagina = tablaPagina?.Rows.Count ?? 0;
            int totalPaginas = _pagerPredef.TotalPages;

            if (totalFilas <= 0)
            {
                _lblCountRowsPredef.Text = "Registros encontrados: 0";
                _lblCountRowsPredef.Visible = true;
                _btnExcelPredef.Visible = false;
                _btnPrevPagePredef.Enabled = false;
                _btnNextPagePredef.Enabled = false;
                return;
            }

            _lblCountRowsPredef.Text =
                $"Registros encontrados: {totalFilas}. " +
                $"Cargando {filasPagina} registros de página {_pagerPredef.CurrentPage + 1} de {totalPaginas}";

            _lblCountRowsPredef.Visible = true;
            _btnExcelPredef.Visible = true;

            _btnPrevPagePredef.Enabled = _pagerPredef.CurrentPage > 0;
            _btnNextPagePredef.Enabled = _pagerPredef.CurrentPage < totalPaginas - 1;
        }

        private void ActualizarPaginacionAdHoc()
        {
            if (_pagerAdHoc == null)
            {
                ResetPaginacionAdHoc();
                return;
            }

            int totalFilas = _pagerAdHoc.TotalRows;
            var tablaPagina = _gridAdHoc.DataSource as DataTable;
            int filasPagina = tablaPagina?.Rows.Count ?? 0;
            int totalPaginas = _pagerAdHoc.TotalPages;

            if (totalFilas <= 0)
            {
                _lblCountRowsAdHoc.Text = "Registros encontrados: 0";
                _lblCountRowsAdHoc.Visible = true;
                _btnExcelAdHoc.Visible = false;
                _btnPrevPageAdHoc.Enabled = false;
                _btnNextPageAdHoc.Enabled = false;
                _btnPrevPageAdHoc.Visible = false;
                _btnNextPageAdHoc.Visible = false;
                return;
            }

            _lblCountRowsAdHoc.Text =
                $"Registros encontrados: {totalFilas}. " +
                $"Cargando {filasPagina} registros de página {_pagerAdHoc.CurrentPage + 1} de {totalPaginas}";

            _lblCountRowsAdHoc.Visible = true;
            _btnExcelAdHoc.Visible = true;
            _btnPrevPageAdHoc.Visible = true;
            _btnNextPageAdHoc.Visible = true;

            _btnPrevPageAdHoc.Enabled = _pagerAdHoc.CurrentPage > 0;
            _btnNextPageAdHoc.Enabled = _pagerAdHoc.CurrentPage < totalPaginas - 1;
        }

        #endregion

        #region Render parámetros / construcción parámetros

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

            FlowLayoutPanel CreateParamBlock(string? labelText, Control input, bool? filterLike)
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

            var preselected = parameter.ValuesRequired ?? new List<string?>();

            int maxWidth = 0;

            foreach (DataRow row in dt.Rows)
            {
                string value = row[parameter.Id].ToString() ?? "";
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
                            value = p.Values?
                                .Where(x => x.Key == numberFromBoolean)?
                                .FirstOrDefault()?
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
                            if (chkLike != null && chkLike.Checked)
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

        #endregion

        #region Exportación Excel

        private void ExportGridWithClosedXml(object? sender, EventArgs e)
        {
            using var cts = new CancellationTokenSource();
            using var loading = new LoadingForm("Exportando datos a Excel ...");

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

                if (sender is Button btn && btn.Tag is ResultTabUI typeExport)
                {
                    // Nombre sugerido del archivo
                    switch (typeExport)
                    {
                        case ResultTabUI.TabInitial:
                            sfd.FileName = $"{((ReportDefinition)_cmbReports.SelectedItem!).Name}_{uniqueIdTime}.xlsx";
                            break;

                        case ResultTabUI.TabSecundary:
                            sfd.FileName = $"ConsultaPersonalizada_{uniqueIdTime}.xlsx";
                            break;
                    }

                    if (sfd.ShowDialog() != DialogResult.OK)
                        return;

                    using var wb = new XLWorkbook();
                    IXLWorksheet? ws = null;

                    // Hoja según pestaña
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

                    DataTable? gridExport = null;

                    switch (typeExport)
                    {
                        case ResultTabUI.TabInitial:
                            gridExport = _pagerPredef?.FullData;
                            break;

                        case ResultTabUI.TabSecundary:
                            gridExport = _pagerAdHoc?.FullData;
                            break;
                    }

                    if (gridExport == null || gridExport.Columns.Count == 0)
                        throw new Exception("No hay datos para exportar.");

                    // Cabeceras
                    for (int col = 0; col < gridExport.Columns.Count; col++)
                    {
                        ws.Cell(1, col + 1).Value = gridExport.Columns[col].ColumnName;
                        ws.Cell(1, col + 1).Style.Font.Bold = true;
                    }

                    // Filas
                    for (int row = 0; row < gridExport.Rows.Count; row++)
                    {
                        for (int col = 0; col < gridExport.Columns.Count; col++)
                        {
                            var value = gridExport.Rows[row][col];
                            string? safeValue = value == null ? "" : value.ToString();
                            ws.Cell(row + 2, col + 1).Value = safeValue?.Trim();
                        }
                    }

                    // Ajuste de columnas
                    ws.Columns().AdjustToContents();
                    foreach (var sheet in wb.Worksheets)
                        sheet.Columns().AdjustToContents();

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
                if (!loading.IsDisposed)
                    loading.Close();
                Enabled = true;
            }
        }


        #endregion

        #region Ver consulta

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

        #region Form de carga (LoadingForm)

        private sealed class LoadingForm : Form
        {
            private readonly CancellationTokenSource? _cts;
            private Panel _buttonsPanel;
            private Button btnCancelar;
            private Label lbl;

            public void ChangeButtonCancelText(string text)
            {
                if (btnCancelar != null)
                {
                    btnCancelar.Text = text;
                    btnCancelar.Refresh();
                }
            }

            public void ChangeVisibilityButton(bool state)
            {
                if (btnCancelar != null)
                {
                    btnCancelar.Visible = state;
                    btnCancelar.Refresh();
                }
            }

            public void ChangeMessage(string message)
            {
                if (lbl != null)
                {
                    lbl.Text = message;
                    lbl.Refresh();
                }
            }

            public void ChangeVisilityMessage(bool state)
            {
                if (lbl != null)
                {
                    lbl.Visible = state;
                    lbl.Refresh();
                }
            }

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

                lbl = new Label
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
                    btnCancelar.Click += BtnCancelar_Click;

                    _buttonsPanel.Controls.Add(btnCancelar);
                    CenterButton();
                    _buttonsPanel.Resize += ButtonsPanel_Resize;
                }
            }

            private void BtnCancelar_Click(object? sender, EventArgs e)
            {
                _cts?.Cancel();
            }

            private void ButtonsPanel_Resize(object? sender, EventArgs e)
            {
                CenterButton();
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

    public sealed class PropertyGrid
    {
        public DataTable? FullData { get; set; }
        public DataGridView Grid { get; set; }
        public ResultTabUI TypeResource { get; set; }
        public ReportDefinition? CurrentReport { get; set; }

        public int PageSize { get; set; } = 500;
        public int CurrentPage { get; private set; } = 0;

        public event EventHandler? PageChanged;

        public int TotalRows => FullData?.Rows.Count ?? 0;

        public int TotalPages =>
            FullData == null || PageSize <= 0
                ? 0
                : (int)Math.Ceiling(TotalRows / (double)PageSize);

        public PropertyGrid(DataTable? fullData, DataGridView grid, ResultTabUI typeResource, ReportDefinition? currentReport)
        {
            FullData = fullData;
            Grid = grid;
            TypeResource = typeResource;
            CurrentReport = currentReport;
        }

        public void ShowPage(int pageIndex)
        {
            if (FullData == null || FullData.Rows.Count == 0)
            {
                Grid.DataSource = null;
                CurrentPage = 0;
                PageChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            if (pageIndex <= 0)
                pageIndex = 0;
            if (pageIndex >= TotalPages)
                pageIndex = TotalPages - 1;

            var rows = FullData.AsEnumerable()
                               .Skip(pageIndex * PageSize)
                               .Take(PageSize);

            DataTable pageTable = rows.Any()
                ? rows.CopyToDataTable()
                : FullData.Clone();

            Grid.SuspendLayout();
            try
            {
                Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                Grid.DataSource = pageTable;
            }
            finally
            {
                Grid.ResumeLayout();
            }

            CurrentPage = pageIndex;
            PageChanged?.Invoke(this, EventArgs.Empty);
        }

        public void ShowFirstPage() => ShowPage(0);
        public void ShowNextPage() => ShowPage(CurrentPage + 1);
        public void ShowPreviousPage() => ShowPage(CurrentPage - 1);
    }

    public class MultiItem
    {
        public string Value { get; set; } = "";
        public string Text { get; set; } = "";

        public override string ToString()
            => $"({Value}) -> {Text}";
    }

    #endregion
}
