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
        private TabPage _tabAdHoc;
        private static readonly Regex RegexParams = new(@"(?<!:):(?!\d)\w+", RegexOptions.Compiled);

        // Paginadores (uno por pestaña)
        private PropertyGrid? _pagerPredef;
        private PropertyGrid? _pagerAdHoc;

        private readonly Button _btnPrevPage = new() { Text = "< Anterior", Name = "_btnPrevPage", AutoSize = true, Visible = false };
        private readonly Button _btnNextPage = new() { Text = "Siguiente >", Name = "_btnNextPage", AutoSize = true, Visible = false };

        private readonly Button _btnPrevPageAdHoc = new() { Text = "< Anterior", Name = "_btnPrevPageAdHoc", AutoSize = true, Visible = false };
        private readonly Button _btnNextPageAdHoc = new() { Text = "Siguiente >", Name = "_btnNextPageAdHoc", AutoSize = true, Visible = false };

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

        private readonly RichTextBox _txtSqlAdHoc = new()
        {
            Dock = DockStyle.Fill,
            Font = new System.Drawing.Font("Consolas", 10f),
            BackColor = System.Drawing.Color.FromArgb(232, 238, 247),   // (#E8EEF7)
            ForeColor = System.Drawing.Color.FromArgb(28, 59, 106),     // (#1C3B6A)
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

        private readonly ToolTip _toolTip = new ToolTip();

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
            _tabAdHoc = new TabPage("SQL avanzada");

            // --- Pestaña de informes predefinidos ---
            _tabPredefinidos.Controls.Add(_grid);
            _tabPredefinidos.Controls.Add(_chkConnections);
            _tabPredefinidos.Controls.Add(_grpParametros);
            _tabPredefinidos.Controls.Add(_topPanel);

            // Panel de paginación para predefinidos
            var paginationPanelPredef = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(5),
                Margin = new Padding(0),
                WrapContents = false
            };

            _btnPrevPage.Click += _btnPrevPage_Click;
            _btnNextPage.Click += _btnNextPage_Click;

            paginationPanelPredef.Controls.Add(_btnNextPage);
            paginationPanelPredef.Controls.Add(_btnPrevPage);

            _tabPredefinidos.Controls.Add(paginationPanelPredef);


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
            _topPanelAdHoc.Controls.Add(ButtonAdHoc);

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

            var paginationPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Padding = new Padding(5),
                Margin = new Padding(0),
                WrapContents = false
            };

            var separationPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.None,
                Height = 20
            };

            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));   // RichTextBox
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // Separación
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));   // Grid
            layoutAdHoc.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // Paginación

            _txtSqlAdHoc.Dock = DockStyle.Fill;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Copiar", null, new EventHandler((_, __) => _txtSqlAdHoc.Copy()));
            menu.Items.Add("Pegar", null, new EventHandler((_, __) => _txtSqlAdHoc.Paste()));
            menu.Items.Add("Cortar", null, new EventHandler((_, __) => _txtSqlAdHoc.Cut()));
            menu.Items.Add("Seleccionar todo", null, new EventHandler((_, __) => _txtSqlAdHoc.SelectAll()));
            _txtSqlAdHoc.ContextMenuStrip = menu;
            _txtSqlAdHoc.Enter += _txtSqlAdHoc_Enter;
            _txtSqlAdHoc.Leave += _txtSqlAdHoc_Leave;

            layoutAdHoc.Controls.Add(_txtSqlAdHoc, 0, 0);

            _gridAdHoc.Dock = DockStyle.Fill;
            layoutAdHoc.Controls.Add(_gridAdHoc, 0, 2);

            layoutAdHoc.Controls.Add(separationPanel, 0, 1);

            _btnPrevPageAdHoc.Click += _btnPrevPageAdHoc_Click;
            _btnNextPageAdHoc.Click += _btnNextPageAdHoc_Click;

            paginationPanel.Controls.Add(_btnNextPageAdHoc);
            paginationPanel.Controls.Add(_btnPrevPageAdHoc);

            layoutAdHoc.Controls.Add(paginationPanel, 0, 3);

            rightPanelAdHoc.Controls.Add(layoutAdHoc);

            _tabAdHoc.Controls.Add(rightPanelAdHoc);
            _tabAdHoc.Controls.Add(_chkConnectionsAdHoc);
            _tabAdHoc.Controls.Add(_topPanelAdHoc);

            _tabControl.TabPages.Add(_tabPredefinidos);
            _tabControl.TabPages.Add(_tabAdHoc);

            Controls.Add(_tabControl);

            // OwnerDraw para colorear pestaña activa
            _tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
            _tabControl.DrawItem += TabControl_DrawItem;
            _tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

            // Botón por defecto inicial: informes predefinidos
            this.AcceptButton = _btnRunReport;

            // Aplicar tema visual
            ApplyTheme();


            Load += MainForm_LoadAsync;
        }

        private void _btnNextPage_Click(object? sender, EventArgs e)
        {
            if (_pagerPredef == null) return;

            _pagerPredef.ShowNextPage();
            PaintControlsTopGrid(_grid, ResultTabUI.TabInitial, _pagerPredef);
        }

        private void _btnPrevPage_Click(object? sender, EventArgs e)
        {
            if (_pagerPredef == null) return;

            _pagerPredef.ShowPreviousPage();
            PaintControlsTopGrid(_grid, ResultTabUI.TabInitial, _pagerPredef);
        }


        private void _btnNextPageAdHoc_Click(object? sender, EventArgs e)
        {
            if (_pagerAdHoc == null) return;

            _pagerAdHoc.ShowNextPage();
            PaintControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary, _pagerAdHoc);
        }

        private void _btnPrevPageAdHoc_Click(object? sender, EventArgs e)
        {
            if (_pagerAdHoc == null) return;

            _pagerAdHoc.ShowPreviousPage();
            PaintControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary, _pagerAdHoc);
        }
        private void _txtSqlAdHoc_Enter(object sender, EventArgs e)
        {
            // Mientras escribo SQL, quiero que Enter sea solo salto de línea
            if (_tabControl.SelectedTab == _tabAdHoc)
            {
                this.AcceptButton = null;
            }
        }

        private void _txtSqlAdHoc_Leave(object sender, EventArgs e)
        {
            // Al salir del editor, vuelvo a activar el botón por defecto de la pestaña
            if (_tabControl.SelectedTab == _tabAdHoc)
            {
                this.AcceptButton = ButtonAdHoc;
            }
        }

        private void _btnClearAdHoc_Click(object? sender, EventArgs e)
        {
            _txtSqlAdHoc.Clear();
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
                    _gridAdHoc.DataSource = null;
                    RemoveControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary);
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
                    MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    RecursiveEnableControlsForm(this, true,true);
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
                    RemoveControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary);
                    _gridAdHoc.DataSource = null;

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
                    {
                        _pagerAdHoc = new PropertyGrid(resultQuery.Data, _gridAdHoc, ResultTabUI.TabSecundary, null);
                        _pagerAdHoc.PageChanged += Pager_PageChanged;
                        _pagerAdHoc.ShowFirstPage();
                        PaintControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary, _pagerAdHoc);
                    }
                }
                catch (OperationCanceledException)
                {
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
                    _gridAdHoc.DataSource = null;
                    MessageBox.Show("Consulta cancelada por el usuario.",
                        "Cancelado",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                finally
                {
                    RecursiveEnableControlsForm(this, true,true);
                    Cursor = Cursors.Default;
                    if (!loadingFormAdHoc.IsDisposed)
                        loadingFormAdHoc.Close();
                    Enabled = true;
                    _btnUnselectAllAdHoc.PerformClick();
                }
            }
        }

        private void Pager_PageChanged(object? sender, EventArgs e)
        {
            if (sender is PropertyGrid propPag)
            {
                if (propPag == _pagerAdHoc)
                {
                    _btnPrevPageAdHoc.Enabled = propPag.CurrentPage==0?true: propPag.CurrentPage > 0;
                    _btnNextPageAdHoc.Enabled = propPag.CurrentPage < propPag.TotalPages - 1;
                    RemoveControlsTopGrid(_gridAdHoc, ResultTabUI.TabSecundary);
                }
                else if (propPag == _pagerPredef)
                {
                    _btnPrevPage.Enabled = propPag.CurrentPage==0?true:propPag.CurrentPage > 0;
                    _btnNextPage.Enabled = propPag.CurrentPage < propPag.TotalPages - 1;
                    RemoveControlsTopGrid(_grid, ResultTabUI.TabInitial);
                }
            }
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

        private void RecursiveEnableControlsForm(Control control, bool changeStated,bool dataLoad=false)
        {
       
            if (control == null)
                return;

          
                control.Enabled = changeStated;
          
                foreach (Control child in control.Controls)
                {

                     bool   skipSetting =
                                control.Name.Contains("_btnPrevPageAdHoc") ||
                                control.Name.Contains("_btnPrevPage") ||
                                control.Name.Contains("_btnNextPageAdHoc") ||
                                control.Name.Contains("_btnNextPage");

                        if (!skipSetting)
                            continue;
                    RecursiveEnableControlsForm(child, changeStated);
                }
            
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
            _topPanel.Controls.Add(_btnVerConsulta);
            _topPanel.Controls.Add(_btnRunReport);

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

        #region Tema visual

        private void ApplyTheme()
        {
            this.BackColor = AppTheme.FormBackColor;

            _topPanel.BackColor = AppTheme.TopPanelBackColor;
            _topPanelAdHoc.BackColor = AppTheme.TopPanelBackColor;
            _grpParametros.BackColor = AppTheme.GroupBoxBackColor;

            ApplyGridTheme(_grid);
            ApplyGridTheme(_gridAdHoc);

            StylePrimaryButton(_btnRunReport);
            StylePrimaryButton(ButtonAdHoc);

            StyleSecondaryButton(_btnSelectAll);
            StyleSecondaryButton(_btnUnselectAll);
            StyleSecondaryButton(_btnSelectAllAdHoc);
            StyleSecondaryButton(_btnUnselectAllAdHoc);
            StyleSecondaryButton(_btnClearAdHoc);
            StyleSecondaryButton(_btnVerConsulta);
            StyleSecondaryButton(_btnPrevPageAdHoc);
            StyleSecondaryButton(_btnNextPageAdHoc);
            StyleSecondaryButton(_btnPrevPage);       
            StyleSecondaryButton(_btnNextPage);       
        }

        private static void ApplyAlternateRowStyle(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;

            // Cabecera
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            // Fondo general
            dgv.BackgroundColor = Color.White;
            dgv.DefaultCellStyle.BackColor = Color.White;

            // Alternancia más marcada
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 244, 255); // azul clarito
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.Black;

            // Bordes y selección
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.GridColor = Color.LightGray;

            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(51, 153, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;

            // Ajustes visuales extra
            dgv.RowHeadersVisible = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }


        private void ApplyGridTheme(DataGridView grid)
        {
            grid.BackgroundColor = AppTheme.GridBackColor;
            grid.DefaultCellStyle.BackColor = AppTheme.GridBackColor;
            grid.DefaultCellStyle.ForeColor = AppTheme.GridForeColor;
            grid.BorderStyle = BorderStyle.FixedSingle;
            grid.GridColor = AppTheme.GridBorderColor;
            grid.DefaultCellStyle.Font = this.Font;

            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.GridHeaderBackColor;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.GridHeaderForeColor;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(grid.Font, FontStyle.Bold);

            grid.AlternatingRowsDefaultCellStyle.BackColor = AppTheme.GridAlternateRowColor;
        }

        private void StylePrimaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = AppTheme.PrimaryButtonBackColor;
            btn.ForeColor = AppTheme.PrimaryButtonForeColor;
            btn.Padding = new Padding(6, 2, 6, 2);
        }

        private void StyleSecondaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = AppTheme.SecondaryButtonBackColor;
            btn.ForeColor = AppTheme.SecondaryButtonForeColor;
            btn.Padding = new Padding(6, 2, 6, 2);
        }

        private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl tab = (TabControl)sender;
            TabPage page = tab.TabPages[e.Index];

            bool isSelected = (e.Index == tab.SelectedIndex);

            Color backColor = isSelected
                ? AppTheme.ActiveTabBackColor
                : AppTheme.InactiveTabBackColor;

            Color borderColor = isSelected
                ? AppTheme.ActiveTabBorderColor
                : AppTheme.InactiveTabBorderColor;

            using (var backBrush = new SolidBrush(backColor))
            using (var borderPen = new Pen(borderColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                e.Graphics.DrawRectangle(borderPen, e.Bounds);
            }

            TextRenderer.DrawText(
                e.Graphics,
                page.Text,
                tab.Font,
                e.Bounds,
                Color.Black,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == _tabPredefinidos)
            {
                this.AcceptButton = _btnRunReport;
            }
            else
            {
                this.AcceptButton = _txtSqlAdHoc.Focused ? null : ButtonAdHoc;
            }
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
            RemoveControlsTopGrid(_grid, ResultTabUI.TabInitial);
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
                RemoveControlsTopGrid(_grid, ResultTabUI.TabInitial);
                _grid.DataSource = null;
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
                    _pagerPredef = new PropertyGrid(resultReport.Data, _grid, ResultTabUI.TabInitial, (ReportDefinition)_cmbReports.SelectedItem);
                    _pagerPredef.PageChanged += Pager_PageChanged;
                    _pagerPredef.ShowFirstPage();
                    PaintControlsTopGrid(_grid, ResultTabUI.TabInitial, _pagerPredef);
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
                MessageBox.Show("Consulta cancelada por el usuario.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                RecursiveEnableControlsForm(this, true, true);
                Cursor = Cursors.Default;
                if (!loading.IsDisposed)
                    loading.Close();
                Enabled = true;
                _btnUnselectAll.PerformClick();
            }
        }

        private void RemoveControlsTopGrid(DataGridView grid, ResultTabUI nameTab)
        {
            if (grid == null || grid.Parent == null)
                return;

            Control parent = grid.Parent;

            if (parent is TableLayoutPanel && parent.Parent != null)
                parent = parent.Parent;

            var listLbl = parent.Controls.OfType<Label>()
                .Where(l => l.Name.Contains("lblCountRows_"))
                .ToList();

            foreach (var itemlbl in listLbl)
                parent.Controls.Remove(itemlbl);

            var listBtn = parent.Controls.OfType<Button>()
                .Where(b => b.Name.Contains("btnExportExcel_"))
                .ToList();

            foreach (var itemButton in listBtn)
                parent.Controls.Remove(itemButton);

            if (nameTab == ResultTabUI.TabSecundary)
            {
                _btnNextPageAdHoc.Visible = false;
                _btnPrevPageAdHoc.Visible = false;
            }
            else if (nameTab == ResultTabUI.TabInitial)
            {
                _btnNextPage.Visible = false;
                _btnPrevPage.Visible = false;
            }

        }

        private void PaintControlsTopGrid(DataGridView? grid, ResultTabUI nameTab, PropertyGrid pager)
        {
            if (grid == null || grid.Parent == null)
                return;

            Control parent = grid.Parent;

            if (parent is TableLayoutPanel && parent.Parent != null)
                parent = parent.Parent;

            System.Drawing.Point gridLocationInParent = parent.PointToClient(
                grid.Parent.PointToScreen(grid.Location));

            string SufixLabel = pager.CurrentReport == null ? nameTab.ToString() : pager.CurrentReport.Id;

            _lblCountRows = new Label
            {
                Name = $"lblCountRows_{SufixLabel}",
                AutoSize = true,
                ForeColor = SystemColors.ControlText,
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Margin = new Padding(4)
            };
            parent.Controls.Add(_lblCountRows);

            int? rowCount = pager.FullData?.Rows.Count;
            if (rowCount > 0)
                _lblCountRows.Text = $"Registros encontrados: {rowCount}. Cargando {((DataTable)grid.DataSource).Rows.Count} registros de página {pager.CurrentPage + 1} de {pager.TotalPages}";
            else
                _lblCountRows.Text = $"Registros encontrados: {rowCount}";

            if (_lblCountRows.Top == 0 && _lblCountRows.Left == 0)
            {
                _lblCountRows.Top = gridLocationInParent.Y - _lblCountRows.Height - 18;
                _lblCountRows.Left = parent.ClientSize.Width - _lblCountRows.Width - 10;
            }
            _lblCountRows.BringToFront();

            string SufixExcel = pager.CurrentReport == null ? nameTab.ToString() : pager.CurrentReport.Id;

            _btnExcel = new Button
            {
                Name = $"btnExportExcel_{SufixExcel}",
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

            if (_btnExcel.Top == 0 && _btnExcel.Left == 0)
            {
                _btnExcel.Top = _lblCountRows.Top - 4;
                _btnExcel.Left = _lblCountRows.Left - _btnExcel.Width - 6;
            }

            _btnExcel.Visible = pager.FullData?.Rows.Count > 0;
            _lblCountRows.BringToFront();
            _btnExcel.BringToFront();

            // Tooltip
            _toolTip.SetToolTip(_btnExcel, "Exportar todos los registros a Excel");

            // Mostrar/ocultar botones de paginación según el tab
            if (pager.FullData?.Rows.Count > 0)
            {
                if (nameTab == ResultTabUI.TabSecundary)
                {
                    _btnNextPageAdHoc.Visible = pager.TotalPages > 1;
                    _btnPrevPageAdHoc.Visible = pager.TotalPages > 1;
                }
                else if (nameTab == ResultTabUI.TabInitial)
                {
                    _btnNextPage.Visible = pager.TotalPages > 1;
                    _btnPrevPage.Visible = pager.TotalPages > 1;
                }
            }

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

        private void SetAllConnectionsChecked(bool isChecked, CheckedListBox chk)
        {
            for (int i = 0; i < chk.Items.Count; i++)
                chk.SetItemChecked(i, isChecked);
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

            // Centrar visualmente en el primer elemento seleccionado
            if (clb.CheckedIndices.Count > 0)
            {
                int firstChecked = clb.CheckedIndices[0];
                clb.SelectedIndex = firstChecked;
                clb.TopIndex = firstChecked;
            }

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

        private void ExportGridWithClosedXml(object? sender, EventArgs e)
        {
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

                    // Exportar SIEMPRE el FullData del paginador (todos los registros)
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

                    for (int col = 0; col < gridExport.Columns.Count; col++)
                    {
                        ws.Cell(1, col + 1).Value = gridExport.Columns[col].ColumnName;
                        ws.Cell(1, col + 1).Style.Font.Bold = true;
                    }

                    for (int row = 0; row < gridExport.Rows.Count; row++)
                    {
                        for (int col = 0; col < gridExport.Columns.Count; col++)
                        {
                            var value = gridExport.Rows[row][col];
                            string? safeValue = value == null ? "" : value.ToString();
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
                RecursiveEnableControlsForm(this, true,true);
                Cursor = Cursors.Default;
                if (!loading.IsDisposed)
                    loading.Close();
                Enabled = true;
            }
        }

        #endregion

        #region Form de carga

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
                    btnCancelar.Click += new EventHandler((_, __) => _cts.Cancel());
                    _buttonsPanel.Controls.Add(btnCancelar);
                    CenterButton();
                    _buttonsPanel.Resize += new EventHandler((_, __) => CenterButton());
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

            DataTable pageTable;

            if (rows.Any())
                pageTable = rows.CopyToDataTable();
            else
                pageTable = FullData.Clone();

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
