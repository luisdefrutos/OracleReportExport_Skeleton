using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OracleReportExport.Domain.Enums;
using OracleReportExport.Domain.Models;

namespace OracleReportExport.Presentation.Desktop
{
    public partial class SaveAdHocReportForm : Form
    {
        private readonly string _sql;
        private readonly List<ParameterRow> _parameterRows = new();

        private TextBox _txtName = default!;
        private TextBox _txtCategory = default!;
        private ComboBox _cboSourceType = default!;
        private Label lblParams = default!;

        // Contenedor con scroll para muchos parámetros
        private Panel _paramContainer = default!;
        // Tabla de parámetros: Label | Combo tipo | Control ejemplo
        private TableLayoutPanel _paramPanel = default!;

        private Button _btnOk = default!;
        private Button _btnCancel = default!;

        // Resultado final: el ReportDefinition construido por el usuario
        public ReportDefinition? Result { get; private set; }

        public SaveAdHocReportForm(string sql)
        {
            _sql = sql ?? throw new ArgumentNullException(nameof(sql));
            MyInitializeComponent();
            LoadParametersFromSql();
        }

        // ----- Inicialización de controles -----

        private void MyInitializeComponent()
        {
            Text = "Guardar informe AdHoc";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(10);
            //MinimumSize = new System.Drawing.Size(650, 0);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            Controls.Add(layout);

            int row = 0;

            // Nombre del informe
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                Text = "Nombre del informe:",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            }, 0, row);

            _txtName = new TextBox
            {
                Dock = DockStyle.Fill
            };
            layout.Controls.Add(_txtName, 1, row);
            row++;

            // Categoría
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                Text = "Categoría:",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            }, 0, row);

            _txtCategory = new TextBox
            {
                Dock = DockStyle.Fill,
                Text = "Ad-Hoc"
            };
            layout.Controls.Add(_txtCategory, 1, row);
            row++;

            // Tipo de informe
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                Text = "Tipo de informe:",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            }, 0, row);

            _cboSourceType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Left,
                Width = 150
            };
            _cboSourceType.Items.AddRange(new object[]
            {
                ReportSourceType.Estacion.ToString(),
                ReportSourceType.Central.ToString(),
                ReportSourceType.Ambos.ToString()
            });
            _cboSourceType.SelectedIndex = 0;
            layout.Controls.Add(_cboSourceType, 1, row);
            row++;

            // Label parámetros
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
             lblParams = new Label
            {
                Text = "Parámetros detectados:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 10, 0, 5)
            };
            layout.Controls.Add(lblParams, 0, row);
            layout.SetColumnSpan(lblParams, 2);
            row++;

            // Panel scrollable para parámetros
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _paramContainer = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true 
                //MinimumSize = new System.Drawing.Size(0, 220)
            };

            // Tabla de parámetros
            _paramPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                Margin= new Padding(0, 15, 0, 0)
            };
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35)); // label
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // combo
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35)); // control ejemplo

            _paramContainer.Controls.Add(_paramPanel);

            layout.Controls.Add(_paramContainer, 0, row);
            layout.SetColumnSpan(_paramContainer, 2);
            row++;

            // Panel de botones
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };

            _btnOk = new Button
            {
                Text = "Guardar",
                DialogResult = DialogResult.OK,
                AutoSize = true
            };
            _btnOk.Click += OnOkClick;
            StylePrimaryButton(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancelar",
                DialogResult = DialogResult.Cancel,
                AutoSize = true
            };

            buttonPanel.Controls.Add(_btnOk);
            buttonPanel.Controls.Add(_btnCancel);

            layout.Controls.Add(buttonPanel, 0, row);
            layout.SetColumnSpan(buttonPanel, 2);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private void StylePrimaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = AppTheme.PrimaryButtonBackColor;
            btn.ForeColor = AppTheme.PrimaryButtonForeColor;
            btn.Padding = new Padding(6, 2, 6, 2);
        }
        // ----- Carga de parámetros detectados en el SQL -----
        private void LoadParametersFromSql()
        {
            var paramNames = ExtractParameterNames(_sql);
            if (paramNames.Count == 0)
            {
                _paramContainer.Visible = false;
                lblParams.Visible = false;
                return;
            }

            _paramPanel.RowCount = paramNames.Count; 

            int rowIndex = 0;
            int increedView = 0;
            foreach (var name in paramNames)
            {
                _paramPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                // Label
                var lbl = new Label
                {
                    Text = Capitalize(name) + ":",
                    AutoSize = true,
                    Anchor = AnchorStyles.Left
                };
                _paramPanel.Controls.Add(lbl, 0, rowIndex);

                // Combo tipo
                var cboType = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Anchor = AnchorStyles.Left,
                    Width = 140
                };
                cboType.Items.AddRange(new object[]
                {
                    "text",
                    "date",
                    "int",
                    "decimal",
                    "bool"
                });
                cboType.SelectedItem = "text";
                _paramPanel.Controls.Add(cboType, 1, rowIndex);

                // Control de ejemplo en la misma fila
                var previewControl = CreatePreviewControl("text");
                previewControl.Anchor = AnchorStyles.Left;
                previewControl.Width = 180;
                _paramPanel.Controls.Add(previewControl, 2, rowIndex);

                var paramRow = new ParameterRow(name, cboType, rowIndex, previewControl);
                _parameterRows.Add(paramRow);

                cboType.SelectedIndexChanged += (_, _) => UpdatePreviewControl(paramRow);

                rowIndex++;
                increedView += previewControl.Height+40;

            }
            MinimumSize=new System.Drawing.Size(650, increedView);
        }


        public static string Capitalize(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            text = text.Trim();

            if (text.Length == 1)
                return text.ToUpper();

            return char.ToUpper(text[0]) + text.Substring(1);
        }
        private static List<string> ExtractParameterNames(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                return new List<string>();

            // Busca :NombreParam (evitando ::)
            var matches = Regex.Matches(sql, @"(?<!:):([A-Za-z_][A-Za-z0-9_]*)");

            return matches
                .Select(m => m.Groups[1].Value)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private sealed class ParameterRow
        {
            public string Name { get; }
            public ComboBox TypeCombo { get; }
            public int RowIndex { get; }
            public Control PreviewControl { get; set; }

            public ParameterRow(string name, ComboBox typeCombo, int rowIndex, Control previewControl)
            {
                Name = name;
                TypeCombo = typeCombo;
                RowIndex = rowIndex;
                PreviewControl = previewControl;
            }
        }

        // ---- Crear control de ejemplo según tipo ----

        private static Control CreatePreviewControl(string type)
        {
            type = (type ?? "text").ToLowerInvariant();

            switch (type)
            {
                case "date":
                    return new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Short
                    };

                case "int":
                    return new NumericUpDown
                    {
                        DecimalPlaces = 0,
                        Maximum = 9999999,
                        Minimum = -9999999
                    };

                case "decimal":
                    return new NumericUpDown
                    {
                        DecimalPlaces = 2,
                        Maximum = 9999999,
                        Minimum = -9999999,
                        Increment = 0.1M
                    };

                case "bool":
                    return new CheckBox
                    {
                        Text = "Sí / No",
                        AutoSize = true
                    };

                case "funcion":
                    return new TextBox
                    {
                        Width = 200
                    };

                case "text":
                default:
                    return new TextBox
                    {
                        Width = 200
                    };
            }
        }

        private void UpdatePreviewControl(ParameterRow row)
        {
            var type = row.TypeCombo.SelectedItem as string ?? "text";

            // Eliminar control anterior de esa celda
            if (row.PreviewControl != null)
            {
                _paramPanel.Controls.Remove(row.PreviewControl);
                row.PreviewControl.Dispose();
            }

            var newControl = CreatePreviewControl(type);
            newControl.Anchor = AnchorStyles.Left;
            newControl.Width = 180;

            row.PreviewControl = newControl;
            _paramPanel.Controls.Add(newControl, 2, row.RowIndex);
        }

        // ----- Click en Guardar -----

        private void OnOkClick(object? sender, EventArgs e)
        {
            // Validaciones básicas
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show("Debe indicar un nombre de informe.", "Guardar informe",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(_txtCategory.Text))
            {
                MessageBox.Show("Debe indicar una categoría.", "Guardar informe",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            // Determinar tipo de informe
            var sourceType = ReportSourceType.Central;
            var selected = _cboSourceType.SelectedItem?.ToString();

            bool asCentral = false;
            bool asEstacion = false;

            if (string.Equals(selected, "Estación", StringComparison.OrdinalIgnoreCase))
            {
                sourceType = ReportSourceType.Estacion;
                asEstacion = true;
            }
            else if (string.Equals(selected, "Central", StringComparison.OrdinalIgnoreCase))
            {
                sourceType = ReportSourceType.Central;
                asCentral = true;
            }
            else // "Ambos"
            {
                sourceType = ReportSourceType.Estacion;
                asCentral = true;
                asEstacion = true;
            }

            // Construir lista de parámetros
            var parameters = new List<ReportParameterDefinition>();

            foreach (var row in _parameterRows)
            {
                var typeString = row.TypeCombo.SelectedItem as string ?? "text";

                parameters.Add(new ReportParameterDefinition
                {
                    Name = row.Name,
                    Label = row.Name,
                    Type = typeString,
                    IsRequired = true,
                    AllowedValues = null,
                    Values = new List<IntCodeItem>(),
                    BusquedaLike = null
                });
            }

            // Construir el ReportDefinition resultado
            Result = new ReportDefinition
            {
                Id = Guid.NewGuid().ToString(),             // si luego quieres ID numérico, lo cambiamos
                Name = _txtName.Text.Trim(),
                Category = _txtCategory.Text.Trim(),
                Description = "Informe guardado desde AdHoc",
                SourceType = sourceType,
                SqlForStations = asEstacion ? _sql : null,
                SqlForCentral = asCentral ? _sql : null,
                Parameters = parameters,
                TableMasterForParameters = Array.Empty<TableMasterParameterDefinition>()
            };
        }
    }
}
