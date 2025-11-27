using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SW2026RibbonAddin; // for ArabicTextUtils

namespace SW2026RibbonAddin.Forms
{
    // partial to allow a Designer partial if you have one
    public partial class FarsiNoteForm : Form
    {
        private readonly bool _enableFormatting;

        private TextBox _txt;
        private ComboBox _fontNames;
        private NumericUpDown _fontSize;
        private Panel _fmtPanel;
        private RadioButton _alignLeft;
        private RadioButton _alignCenter;
        private RadioButton _alignRight;
        private Button _ok;
        private Button _cancel;

        // Default: used for "new note" (formatting enabled)
        public FarsiNoteForm() : this(true)
        {
        }

        // enableFormatting = true  → New-note mode (font/size/alignment row visible)
        // enableFormatting = false → Edit-note mode (text only, formatting row hidden)
        public FarsiNoteForm(bool enableFormatting)
        {
            _enableFormatting = enableFormatting;

            Text = "یادداشت فارسی";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 360);
            Font = new Font("Segoe UI", 9F);

            // Caption
            var lbl = new Label
            {
                Text = "متن (Farsi text):",
                AutoSize = true,
                Left = 12,
                Top = 12
            };
            Controls.Add(lbl);

            // Formatting row: Font + Size + Alignment (R/C/L)
            _fmtPanel = new Panel
            {
                Left = 12,
                Top = 32,
                Width = 496,
                Height = 30
            };
            Controls.Add(_fmtPanel);

            var fontLbl = new Label
            {
                Left = 0,
                Top = 7,
                Width = 40,
                Text = "Font:"
            };
            _fmtPanel.Controls.Add(fontLbl);

            _fontNames = new ComboBox
            {
                Left = 45,
                Top = 3,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var preferred = new[]
            {
                "B Nazanin", "B Zar", "IranNastaliq", "Vazirmatn",
                "Tahoma", "Segoe UI", "Arial Unicode MS"
            };

            var installed = FontFamily.Families.Select(f => f.Name).ToList();
            foreach (var f in preferred)
            {
                if (installed.Contains(f))
                    _fontNames.Items.Add(f);
            }
            if (_fontNames.Items.Count == 0)
            {
                foreach (var f in installed)
                    _fontNames.Items.Add(f);
            }
            if (_fontNames.Items.Count > 0)
                _fontNames.SelectedIndex = 0;

            _fmtPanel.Controls.Add(_fontNames);

            var sizeLbl = new Label
            {
                Left = 255,
                Top = 7,
                Width = 40,
                Text = "Size:"
            };
            _fmtPanel.Controls.Add(sizeLbl);

            _fontSize = new NumericUpDown
            {
                Left = 300,
                Top = 4,
                Width = 55,
                Minimum = 6,
                Maximum = 200,
                Increment = 1,
                Value = 12
            };
            _fmtPanel.Controls.Add(_fontSize);

            // Alignment buttons (R / C / L) for new-note mode
            _alignRight = new RadioButton
            {
                Left = 370,
                Top = 6,
                Width = 40,
                Text = "R",
                Checked = true
            };
            _fmtPanel.Controls.Add(_alignRight);

            _alignCenter = new RadioButton
            {
                Left = 410,
                Top = 6,
                Width = 40,
                Text = "C"
            };
            _fmtPanel.Controls.Add(_alignCenter);

            _alignLeft = new RadioButton
            {
                Left = 450,
                Top = 6,
                Width = 40,
                Text = "L"
            };
            _fmtPanel.Controls.Add(_alignLeft);

            // Main text area
            _txt = new TextBox
            {
                Left = 12,
                Width = 496,
                Height = 210,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                RightToLeft = RightToLeft.Yes
            };

            // Layout depending on mode
            if (_enableFormatting)
            {
                _fmtPanel.Visible = true;
                _fmtPanel.Enabled = true;
                _txt.Top = _fmtPanel.Bottom + 8;
            }
            else
            {
                _fmtPanel.Visible = false;
                _fmtPanel.Enabled = false;
                _txt.Top = 32; // directly under the label
            }

            Controls.Add(_txt);

            // OK / Cancel buttons
            _ok = new Button
            {
                Text = "OK",
                Left = 360,
                Top = 315,
                Width = 70,
                DialogResult = DialogResult.OK
            };
            Controls.Add(_ok);

            _cancel = new Button
            {
                Text = "Cancel",
                Left = 438,
                Top = 315,
                Width = 70,
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(_cancel);

            AcceptButton = _ok;
            CancelButton = _cancel;
        }

        // Text in the editor (auto-normalizes anything coming from SolidWorks)
        public string NoteText
        {
            get => _txt.Text;
            set
            {
                var src = value ?? string.Empty;
                _txt.Text = ArabicTextUtils.FromSolidWorks(src);
            }
        }

        // RTL markers option removed; always false, property kept for compatibility.
        public bool UseRtlMarkers
        {
            get => false;
            set { /* ignored */ }
        }

        // Shaping (fix disconnected letters) always on; property kept for compatibility.
        public bool InsertJoiners
        {
            get => true;
            set { /* ignored */ }
        }

        public string FontFamilyName
        {
            get
            {
                var s = _fontNames.SelectedItem?.ToString();
                return string.IsNullOrWhiteSpace(s) ? "Tahoma" : s;
            }
            set { SelectFont(value); }
        }

        // Back‑compat alias
        public string SelectedFontName
        {
            get => FontFamilyName;
            set => FontFamilyName = value;
        }

        public double FontSizePoints
        {
            get => (double)_fontSize.Value;
            set
            {
                try
                {
                    decimal v = (decimal)value;
                    if (v < _fontSize.Minimum) v = _fontSize.Minimum;
                    if (v > _fontSize.Maximum) v = _fontSize.Maximum;
                    _fontSize.Value = v;
                }
                catch
                {
                    // ignore invalid inputs
                }
            }
        }

        // Alignment for new-note creation (Right by default)
        public HorizontalAlignment Alignment
        {
            get
            {
                if (_alignLeft != null && _alignLeft.Checked) return HorizontalAlignment.Left;
                if (_alignCenter != null && _alignCenter.Checked) return HorizontalAlignment.Center;
                return HorizontalAlignment.Right;
            }
            set
            {
                if (_alignLeft == null || _alignCenter == null || _alignRight == null) return;

                _alignLeft.Checked = (value == HorizontalAlignment.Left);
                _alignCenter.Checked = (value == HorizontalAlignment.Center);
                _alignRight.Checked = (value == HorizontalAlignment.Right);
            }
        }

        private void SelectFont(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                if (_fontNames.Items.Count > 0)
                    _fontNames.SelectedIndex = 0;
                return;
            }

            int idx = -1;
            for (int i = 0; i < _fontNames.Items.Count; i++)
            {
                if (string.Equals(_fontNames.Items[i]?.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    idx = i;
                    break;
                }
            }

            if (idx >= 0)
                _fontNames.SelectedIndex = idx;
            else if (_fontNames.Items.Count > 0)
                _fontNames.SelectedIndex = 0;
        }
    }
}
