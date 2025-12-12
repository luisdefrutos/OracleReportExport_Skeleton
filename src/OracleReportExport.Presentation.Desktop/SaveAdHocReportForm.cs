using OracleReportExport.Application.Models;
using OracleReportExport.Domain.Enums;
using OracleReportExport.Domain.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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
        public bool ErrorParameter { get; private set; } = false;

        // Valores introducidos por el usuario para esta ejecución AdHoc
        public Dictionary<string, object?> RuntimeParameterValues { get; } = new();

        public SaveAdHocReportForm(string sql, List<ConnectionInfo> listConnectionSelected)
        {
            _sql = sql ?? throw new ArgumentNullException(nameof(sql));
            MyInitializeComponent();
            LoadParametersFromSql();
            selectedTypeConnection(listConnectionSelected);



        }

        private void selectedTypeConnection(List<ConnectionInfo> listConnectionSelected)
        {
            bool bothConnection = listConnectionSelected.Any(x => x.Type == ReportSourceType.Central.ToString()) &&
                                    listConnectionSelected.Any(x => x.Type == ReportSourceType.Estacion.ToString());

            bool tieneCentral = listConnectionSelected
                             .Any(x => x.Type.Equals(ReportSourceType.Central.ToString(), StringComparison.OrdinalIgnoreCase));

            bool tieneEstacion = listConnectionSelected
                             .Any(x => x.Type.Equals(ReportSourceType.Estacion.ToString(), StringComparison.OrdinalIgnoreCase));

            bool soloCentral = tieneCentral && !tieneEstacion;
            bool soloEstacion = tieneEstacion && !tieneCentral;
            bool hayAmbos = tieneCentral && tieneEstacion;
            if (hayAmbos)
                _cboSourceType.SelectedItem = ReportSourceType.Ambos.ToString();
            else if (soloCentral)
                _cboSourceType.SelectedItem = ReportSourceType.Central.ToString();
            else
                _cboSourceType.SelectedItem = ReportSourceType.Estacion.ToString();
        }

        // ----- Inicialización de controles -----

        private void MyInitializeComponent()
        {
            Text = "Datos necesarios para poder ejecutar informe AdHoc";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(10);

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
            };

            // Tabla de parámetros
            _paramPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                Margin = new Padding(0, 15, 0, 0)
            };
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30)); // label
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // combo
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45)); // control ejemplo
            _paramPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); // (no se usa ahora)

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
                Text = "Ejecutar",
                DialogResult = DialogResult.OK,
                AutoSize = true
            };
            _btnOk.Click += OnOkClick;
            StylePrimaryButton(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancelar",
                DialogResult = DialogResult.Cancel,
                Height=30
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
            var result = CheckParamRepeat(paramNames);
            if (result)
            {
                this.ErrorParameter = true;
                return;
            }

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
                    Width = 110
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
                var previewControl = CreatePreviewControl("text", name);
                previewControl.Anchor = AnchorStyles.Left;
                previewControl.Width = 180;
                _paramPanel.Controls.Add(previewControl, 2, rowIndex);

                var paramRow = new ParameterRow(name, cboType, rowIndex, previewControl);
                _parameterRows.Add(paramRow);

                cboType.SelectedIndexChanged += (_, _) => UpdatePreviewControl(paramRow);

                rowIndex++;
                increedView += previewControl.Height + 40;
            }

            MinimumSize = new System.Drawing.Size(650, increedView);
        }

        private bool CheckParamRepeat(List<string> paramNames)
        {
            var repetidos = paramNames
               .Select(p => p.Trim())
               .GroupBy(p => p, StringComparer.OrdinalIgnoreCase)
               .Where(g => g.Count() > 1)
               .Select(g => new
               {
                   Nombre = g.Key,
                   Cantidad = g.Count()
               })
               .ToList();

            if (repetidos.Count > 0)
            {
                string mensaje = "Los siguientes parámetros están repetidos en la consulta SQL:\n\n";
                foreach (var item in repetidos)
                {
                    mensaje += $"- {item.Nombre} (repetido {item.Cantidad} veces)\n";
                }
                MessageBox.Show(mensaje, "Parámetros repetidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return repetidos.Count > 0;
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

            var matches = Regex.Matches(sql, @"(?<!:)(?<!\w):\s*([A-Za-z_][A-Za-z0-9_]*)");
            return matches
                .Select(m => m.Groups[1].Value)  // el nombre limpio
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

        private static Control CreatePreviewControl(string type, string parameterName)
        {
            type = (type ?? "text").ToLowerInvariant();

            switch (type)
            {
                case "date":
                    return new DateTimePicker
                    {
                        Format = DateTimePickerFormat.Short,
                        Name = $"{parameterName}_{type}"
                    };

                case "int":
                    return new NumericUpDown
                    {
                        DecimalPlaces = 0,
                        Maximum = 9999999,
                        Minimum = -9999999,
                        Name = $"{parameterName}_{type}"
                    };

                case "decimal":
                    return new NumericUpDown
                    {
                        DecimalPlaces = 2,
                        Maximum = 9999999,
                        Minimum = -9999999,
                        Increment = 0.1M,
                        Name = $"{parameterName}_{type}"
                    };

                case "bool":
                    return new CheckBox
                    {
                        Text = "Sí / No",
                        AutoSize = true,
                        Name = $"{parameterName}_{type}"
                    };

                case "funcion":
                    return new TextBox
                    {
                        Width = 200,
                        Name = $"{parameterName}_{type}"
                    };

                case "text":
                    // TextBox + Check "Búsqueda LIKE" en la misma celda
                    var panel = new FlowLayoutPanel
                    {
                        AutoSize = true,
                        FlowDirection = FlowDirection.LeftToRight,
                        WrapContents = false,
                        Margin = new Padding(0, 0, 0, 0),
                        Name = $"{parameterName}_{type}_panel"
                    };

                    var txt = new TextBox
                    {
                        Width = 150,
                        Name = $"{parameterName}_{type}_text"
                    };

                    var chk = new CheckBox
                    {
                        Text = "Búsqueda 'LIKE'",
                        AutoSize = true,
                        Margin = new Padding(8, 3, 0, 3),
                        Name = $"{parameterName}_{type}_like"
                    };

                    panel.Controls.Add(txt);
                    panel.Controls.Add(chk);

                    return panel;

                default:
                    return new TextBox
                    {
                        Width = 200,
                        Name = $"{parameterName}_{type}"
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

            var newControl = CreatePreviewControl(type, row.Name);
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

            if (string.Equals(selected, ReportSourceType.Estacion.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                sourceType = ReportSourceType.Estacion;
                asEstacion = true;
            }
            else if (string.Equals(selected, ReportSourceType.Central.ToString(), StringComparison.OrdinalIgnoreCase))
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

            // Construir lista de parámetros + valores runtime
            var parameters = new List<ReportParameterDefinition>();
            RuntimeParameterValues.Clear();

            foreach (var row in _parameterRows)
            {
                var typeStringRaw = row.TypeCombo.SelectedItem as string ?? "text";
                var typeString = typeStringRaw.ToLowerInvariant();
                object? value = null;
                bool? busquedaLike = null;

                switch (typeString)
                {
                    case "date":
                        if (row.PreviewControl is DateTimePicker dtp)
                        {
                            var dt = dtp.Value;
                            var nameUpper = row.Name.ToUpperInvariant();

                            // Tratamiento genérico de fechas (sin nombres a fuego exactos)
                            if (nameUpper.Contains("DESDE") ||
                                nameUpper.Contains("INICIO") ||
                                nameUpper.Contains("FROM") ||
                                nameUpper.Contains("START"))
                            {
                                value = dt.Date; // 00:00:00
                            }
                            else if (nameUpper.Contains("HASTA") ||
                                     nameUpper.Contains("FIN") ||
                                     nameUpper.Contains("TO") ||
                                     nameUpper.Contains("END"))
                            {
                                value = dt.Date.AddHours(23).AddMinutes(59).AddSeconds(59); // 23:59:59
                            }
                            else
                            {
                                value = dt; // Fecha/hora tal cual
                            }
                        }
                        break;

                    case "int":
                        if (row.PreviewControl is NumericUpDown nudInt)
                            value = Convert.ToInt32(nudInt.Value);
                        break;

                    case "decimal":
                        if (row.PreviewControl is NumericUpDown nudDec)
                            value = nudDec.Value;
                        break;

                    case "bool":
                        if (row.PreviewControl is CheckBox chkBool)
                            value = chkBool.Checked;
                        break;

                    case "funcion":
                        if (row.PreviewControl is TextBox txtFuncion)
                        {
                            var raw = txtFuncion.Text?.Trim();
                            value = string.IsNullOrWhiteSpace(raw) ? null : raw;
                        }
                        break;

                    case "text":
                    default:
                        // Nuestro diseño para "text" es un FlowLayoutPanel con TextBox + CheckBox LIKE
                        if (row.PreviewControl is FlowLayoutPanel flp)
                        {
                            var txt = flp.Controls.OfType<TextBox>().FirstOrDefault();
                            var chkLike = flp.Controls.OfType<CheckBox>().FirstOrDefault();

                            var raw = txt?.Text?.Trim();
                            value = string.IsNullOrWhiteSpace(raw) ? null : raw;

                            if (chkLike != null)
                            {
                                busquedaLike = chkLike.Checked;
                                if (busquedaLike.Value && value != null)
                                {
                                    value = $"%{value}%";
                                }
                            }
                        }
                        else if (row.PreviewControl is TextBox txtSolo)
                        {
                            var raw = txtSolo.Text?.Trim();
                            value = string.IsNullOrWhiteSpace(raw) ? null : raw;
                        }
                        break;
                }

                // (Opcional) Si quieres exigir valor para todos:
                // if (value == null)
                // {
                //     MessageBox.Show($"Debe indicar un valor para el parámetro \"{row.Name}\".",
                //         "Parámetros incompletos",
                //         MessageBoxButtons.OK,
                //         MessageBoxIcon.Warning);
                //     DialogResult = DialogResult.None;
                //     return;
                // }

                parameters.Add(new ReportParameterDefinition
                {
                    Name = row.Name,
                    Label = Capitalize(row.Name),
                    Type = typeString,
                    IsRequired = true,
                    AllowedValues = null,
                    Values = new List<IntCodeItem>(),
                    BusquedaLike = busquedaLike
                });

                RuntimeParameterValues[row.Name] = value;
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



