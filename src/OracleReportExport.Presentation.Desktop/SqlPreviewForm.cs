using OracleReportExport.Domain.Enums;
using OracleReportExport.Domain.Models;
using System;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OracleReportExport.Presentation.Desktop
{
    public partial class SqlPreviewForm : Form
    {
        private readonly ReportDefinition _report;

        private readonly RichTextBox _txtSql;
        private readonly Button _btnDescargar;
        private readonly Button _btnCopiar;

        // Para ocultar el caret (barra de texto) y que parezca visor
        [DllImport("user32.dll")]
        private static extern bool HideCaret(IntPtr hWnd);

        public SqlPreviewForm(ReportDefinition report, DbConnection cn)
        {
            InitializeComponent();
            _report = report ?? throw new ArgumentNullException(nameof(report));
            Text = "Consulta SQL del informe";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(900, 600);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            // -------- RICHTEXTBOX SOLO LECTURA --------
            _txtSql = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                Font = new Font("Consolas", 10f),
                //ForeColor = Color.FromArgb(40, 40, 40),
                //BackColor = Color.White,
                BackColor = Color.FromArgb(232, 238, 247),   // (#E8EEF7)
                ForeColor = Color.FromArgb(28, 59, 106),     // (#1C3B6A)
                ScrollBars = RichTextBoxScrollBars.Both,
                BorderStyle = BorderStyle.FixedSingle,
                TabStop = false,
                Cursor = Cursors.Arrow
            };

            // Ocultar caret para que parezca visor
            _txtSql.GotFocus += (s, e) => HideCaret(_txtSql.Handle);
            _txtSql.MouseDown += (s, e) => HideCaret(_txtSql.Handle);
            _txtSql.MouseUp += (s, e) => HideCaret(_txtSql.Handle);
            _txtSql.KeyDown += (s, e) => HideCaret(_txtSql.Handle);

            // -------- Panel abajo con botones --------
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 45
            };
            _btnDescargar = new Button
            {
                Text = "Descargar SQL",
                Anchor = AnchorStyles.Right | AnchorStyles.Top,
                AutoSize = true
            };
            _btnDescargar.Click += BtnDescargar_Click;

            _btnCopiar = new Button
            {
                Text = "Copiar",
                Anchor = AnchorStyles.Right | AnchorStyles.Top,
                AutoSize = true
            };
            _btnCopiar.Click += BtnCopiar_Click;

            var btnCerrar = new Button
            {
                Text = "Cerrar",
                Anchor = AnchorStyles.Right | AnchorStyles.Top,
                AutoSize = true
            };
            btnCerrar.Click += (s, e) => Close();

            bottomPanel.Controls.Add(_btnDescargar);
            bottomPanel.Controls.Add(_btnCopiar);
            bottomPanel.Controls.Add(btnCerrar);

            const int padding = 20;
            btnCerrar.Top = _btnCopiar.Top = _btnDescargar.Top = 10;

            void LayoutButtons()
            {
                btnCerrar.Left = bottomPanel.Width - btnCerrar.Width - padding;
                _btnCopiar.Left = btnCerrar.Left - _btnCopiar.Width - padding;
                _btnDescargar.Left = _btnCopiar.Left - _btnDescargar.Width - padding;
            }

            bottomPanel.Resize += (s, e) => LayoutButtons();
            LayoutButtons();

            Controls.Add(_txtSql);
            Controls.Add(bottomPanel);

            // -------- CARGAR SQL --------
            string sql = _report.SourceType == ReportSourceType.Estacion
                ? _report.SqlForStations ?? string.Empty
                : _report.SqlForCentral ?? string.Empty;

            _txtSql.Text = sql.Trim();
            _txtSql.SelectionStart = 0;
            _txtSql.SelectionLength = 0;
        }

        private void BtnDescargar_Click(object? sender, EventArgs e)
        {
            using var sfd = new SaveFileDialog
            {
                Filter = "Fichero SQL (*.sql)|*.sql",
                FileName = $"Consulta_{DateTime.Now:yyyyMMdd_HHmmss}.sql"
            };

            if (sfd.ShowDialog(this) != DialogResult.OK)
                return;

            File.WriteAllText(sfd.FileName, _txtSql.Text, Encoding.UTF8);

            MessageBox.Show("Consulta guardada correctamente.",
                "Descarga", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void BtnCopiar_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtSql.Text))
                return;

            Clipboard.SetText(_txtSql.Text);
            MessageBox.Show("Consulta copiada al portapapeles.",
                "Copiar", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

