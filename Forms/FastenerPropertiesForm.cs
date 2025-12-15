using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Forms
{
    internal sealed class FastenerPropertiesForm : Form
    {
        private readonly FastenerInitialValues _values;

        private ComboBox _family;
        private TextBox _partNumber;
        private ComboBox _standard;
        private TextBox _nominalDiameter;
        private TextBox _length;
        private ComboBox _strengthClass;
        private Label _typeLabel;
        private ComboBox _typeCombo;
        private TextBox _outerDiameter;
        private TextBox _thickness;
        private TextBox _height;
        private TextBox _description;
        private TextBox _material;

        private Button _ok;
        private Button _cancel;

        public FastenerPropertiesForm(FastenerInitialValues values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));

            Text = "Set fastener properties";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 360);

            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            InitializeLayout();
            LoadFromValues();
        }

        private void InitializeLayout()
        {
            int leftLabel = 14;
            int leftInput = 150;
            int top = 14;
            int row = 0;
            int rowH = 24;

            // Fastener family
            var lblFamily = new Label
            {
                Text = "Fastener family:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblFamily);

            _family = new ComboBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _family.Items.AddRange(new object[] { "Bolt", "Washer", "Nut" });
            _family.SelectedIndexChanged += (_, __) => UpdateTypeUI();
            Controls.Add(_family);

            row++;

            // PartNumber
            var lblPN = new Label
            {
                Text = "Part number:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblPN);

            _partNumber = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 200
            };
            Controls.Add(_partNumber);

            row++;

            // Standard
            var lblStd = new Label
            {
                Text = "Standard:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblStd);

            _standard = new ComboBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            _standard.Items.AddRange(new object[]
            {
                "ISO4014", "ISO4017", "ISO4762",
                "ISO7045", "ISO7046-1", "ISO2009", "ISO2010",
                "ISO7089", "ISO7090", "ISO7092", "ISO7093", "ISO7094",
                "ISO4032", "ISO4033", "ISO8673"
            });
            Controls.Add(_standard);

            row++;

            // Nominal diameter (M)
            var lblDia = new Label
            {
                Text = "Nominal diameter (M):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblDia);

            _nominalDiameter = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 80
            };
            Controls.Add(_nominalDiameter);

            row++;

            // Length (bolts)
            var lblLen = new Label
            {
                Text = "Length (mm):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblLen);

            _length = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 80
            };
            Controls.Add(_length);

            row++;

            // Strength class
            var lblStrength = new Label
            {
                Text = "Strength class:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblStrength);

            _strengthClass = new ComboBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            _strengthClass.Items.AddRange(new object[] { "4.6", "5.8", "8", "8.8", "10", "10.9", "12.9" });
            Controls.Add(_strengthClass);

            row++;

            // Type / WasherType / NutType
            _typeLabel = new Label
            {
                Text = "Type:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(_typeLabel);

            _typeCombo = new ComboBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            Controls.Add(_typeCombo);

            row++;

            // Outer diameter (washers)
            var lblOD = new Label
            {
                Text = "Outer diameter (mm):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblOD);

            _outerDiameter = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 80
            };
            Controls.Add(_outerDiameter);

            row++;

            // Thickness (washers)
            var lblThk = new Label
            {
                Text = "Thickness (mm):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblThk);

            _thickness = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 80
            };
            Controls.Add(_thickness);

            row++;

            // Height (nuts - optional)
            var lblHeight = new Label
            {
                Text = "Height (mm, optional):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblHeight);

            _height = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 80
            };
            Controls.Add(_height);

            row++;

            // Description (auto-built suggestion, editable)
            var lblDesc = new Label
            {
                Text = "Description:",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblDesc);

            _description = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 330
            };
            Controls.Add(_description);

            row++;

            // Material (optional)
            var lblMat = new Label
            {
                Text = "Material (optional):",
                AutoSize = true,
                Left = leftLabel,
                Top = top + row * rowH + 3
            };
            Controls.Add(lblMat);

            _material = new TextBox
            {
                Left = leftInput,
                Top = top + row * rowH,
                Width = 330
            };
            Controls.Add(_material);

            // OK / Cancel buttons
            _ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = ClientSize.Width - 180,
                Top = ClientSize.Height - 40,
                Width = 70,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            _ok.Click += OkOnClick;
            Controls.Add(_ok);

            _cancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = ClientSize.Width - 100,
                Top = ClientSize.Height - 40,
                Width = 70,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };
            Controls.Add(_cancel);

            AcceptButton = _ok;
            CancelButton = _cancel;
        }

        private void LoadFromValues()
        {
            var family = _values.Family;
            if (string.IsNullOrWhiteSpace(family))
                family = "Bolt";

            if (family.Equals("Washer", StringComparison.OrdinalIgnoreCase))
                _family.SelectedItem = "Washer";
            else if (family.Equals("Nut", StringComparison.OrdinalIgnoreCase))
                _family.SelectedItem = "Nut";
            else
                _family.SelectedItem = "Bolt";

            _partNumber.Text = _values.PartNumber ?? string.Empty;
            _standard.Text = _values.Standard ?? string.Empty;
            _nominalDiameter.Text = _values.NominalDiameter ?? string.Empty;
            _length.Text = _values.Length ?? string.Empty;
            _strengthClass.Text = _values.StrengthClass ?? string.Empty;

            _outerDiameter.Text = _values.OuterDiameter ?? string.Empty;
            _thickness.Text = _values.Thickness ?? string.Empty;
            _height.Text = _values.Height ?? string.Empty;
            _description.Text = _values.Description ?? string.Empty;
            _material.Text = _values.Material ?? string.Empty;

            UpdateTypeUI();

            // Preselect type by family
            if (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
                _typeCombo.Text = _values.BoltType ?? string.Empty;
            else if (family.Equals("Washer", StringComparison.OrdinalIgnoreCase))
                _typeCombo.Text = _values.WasherType ?? string.Empty;
            else if (family.Equals("Nut", StringComparison.OrdinalIgnoreCase))
                _typeCombo.Text = _values.NutType ?? string.Empty;
        }

        private void OkOnClick(object sender, EventArgs e)
        {
            if (!ValidateAndStore())
                return;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void UpdateTypeUI()
        {
            var fam = (_family.SelectedItem as string) ?? "Bolt";
            fam = fam.Trim();

            _typeCombo.Items.Clear();

            if (fam.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
            {
                _typeLabel.Text = "Bolt type:";
                _typeCombo.Items.AddRange(new object[]
                {
                    "HexBolt", "SocketHeadCap", "PanHead", "Countersunk"
                });

                // Show: length + strength
                _length.Visible = true;
                _strengthClass.Visible = true;
            }
            else if (fam.Equals("Washer", StringComparison.OrdinalIgnoreCase))
            {
                _typeLabel.Text = "Washer type:";
                _typeCombo.Items.AddRange(new object[]
                {
                    "PlainWasher", "SpringWasher"
                });

                // Washers: no length, usually no strength class
                _length.Visible = false;
                _strengthClass.Visible = false;
            }
            else
            {
                _typeLabel.Text = "Nut type:";
                _typeCombo.Items.AddRange(new object[]
                {
                    "HexNut", "LockNut"
                });

                // Nuts: strength class optional, keep visible
                _length.Visible = false;
                _strengthClass.Visible = true;
            }

            // OuterDiameter / Thickness / Height visibility
            bool isWasher = fam.Equals("Washer", StringComparison.OrdinalIgnoreCase);
            bool isNut = fam.Equals("Nut", StringComparison.OrdinalIgnoreCase);

            _outerDiameter.Visible = isWasher;
            _thickness.Visible = isWasher;
            _height.Visible = isNut;
        }

        private bool ValidateAndStore()
        {
            var fam = (_family.SelectedItem as string) ?? "Bolt";
            fam = fam.Trim();

            // PartNumber
            if (string.IsNullOrWhiteSpace(_partNumber.Text))
            {
                MessageBox.Show("PartNumber must not be empty.", "Fastener properties");
                _partNumber.Focus();
                return false;
            }

            // Standard
            if (string.IsNullOrWhiteSpace(_standard.Text))
            {
                MessageBox.Show("Standard must be selected.", "Fastener properties");
                _standard.Focus();
                return false;
            }

            // Nominal diameter > 0
            if (!TryValidatePositive(_nominalDiameter.Text, "Nominal diameter (M) must be a positive number.", out var dia))
            {
                _nominalDiameter.Focus();
                return false;
            }

            // Bolts: Length > 0
            if (fam.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
            {
                if (!TryValidatePositive(_length.Text, "Length must be a positive number for bolts.", out var len))
                {
                    _length.Focus();
                    return false;
                }
            }

            // Washers: OD and thickness > 0
            if (fam.Equals("Washer", StringComparison.OrdinalIgnoreCase))
            {
                if (!TryValidatePositive(_outerDiameter.Text, "Outer diameter must be a positive number for washers.", out var od))
                {
                    _outerDiameter.Focus();
                    return false;
                }

                if (!TryValidatePositive(_thickness.Text, "Thickness must be a positive number for washers.", out var thk))
                {
                    _thickness.Focus();
                    return false;
                }
            }

            // Nuts: diameter already checked; height optional, only check if filled
            if (fam.Equals("Nut", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(_height.Text))
            {
                if (!TryValidatePositive(_height.Text, "Height must be a positive number if specified.", out var h))
                {
                    _height.Focus();
                    return false;
                }
            }

            // Store back into _values (normalized numeric strings)
            _values.Family = fam;
            _values.PartNumber = _partNumber.Text.Trim();
            _values.Standard = _standard.Text.Trim();
            _values.NominalDiameter = NormalizeNumeric(_nominalDiameter.Text);
            _values.Length = fam.Equals("Bolt", StringComparison.OrdinalIgnoreCase)
                ? NormalizeNumeric(_length.Text)
                : string.Empty;

            _values.StrengthClass = fam.Equals("Washer", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : _strengthClass.Text.Trim();

            if (fam.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
            {
                _values.BoltType = _typeCombo.Text.Trim();
                _values.WasherType = null;
                _values.NutType = null;
            }
            else if (fam.Equals("Washer", StringComparison.OrdinalIgnoreCase))
            {
                _values.BoltType = null;
                _values.WasherType = _typeCombo.Text.Trim();
                _values.NutType = null;
            }
            else
            {
                _values.BoltType = null;
                _values.WasherType = null;
                _values.NutType = _typeCombo.Text.Trim();
            }

            _values.OuterDiameter = fam.Equals("Washer", StringComparison.OrdinalIgnoreCase)
                ? NormalizeNumeric(_outerDiameter.Text)
                : string.Empty;

            _values.Thickness = fam.Equals("Washer", StringComparison.OrdinalIgnoreCase)
                ? NormalizeNumeric(_thickness.Text)
                : string.Empty;

            _values.Height = fam.Equals("Nut", StringComparison.OrdinalIgnoreCase)
                ? NormalizeNumeric(_height.Text)
                : string.Empty;

            // Description: if user left blank, rebuild suggestion
            var desc = _description.Text.Trim();
            if (string.IsNullOrWhiteSpace(desc))
            {
                desc = FastenerPropertyHelper.BuildDescription(_values);
            }
            _values.Description = desc;

            _values.Material = _material.Text.Trim();

            return true;
        }

        private static bool TryValidatePositive(string text, string errorMessage, out double value)
        {
            value = 0;
            var raw = text?.Trim();
            if (string.IsNullOrEmpty(raw))
                return false;

            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                !double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                MessageBox.Show(errorMessage, "Fastener properties");
                return false;
            }

            if (value <= 0)
            {
                MessageBox.Show(errorMessage, "Fastener properties");
                return false;
            }

            return true;
        }

        private static string NormalizeNumeric(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var raw = text.Trim();
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                return d.ToString("0.###", CultureInfo.InvariantCulture);
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out d))
                return d.ToString("0.###", CultureInfo.InvariantCulture);

            // fallback
            return raw;
        }
    }
}
