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
        private TextBox _txt;
        private CheckBox _rtl;
        private CheckBox _zwj;
        private ComboBox _fontNames;
        private NumericUpDown _fontSize;
        private Button _ok;
        private Button _cancel;

        public FarsiNoteForm()
        {
            Text = "افزودن یادداشت فارسی";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 340);
            Font = new Font("Segoe UI", 9F);

            var lbl = new Label
            {
                Text = "متن (Farsi text):",
                AutoSize = true,
                Left = 12,
                Top = 12
            };
            Controls.Add(lbl);

            _txt = new TextBox
            {
                Left = 12,
                Top = 32,
                Width = 496,
                Height = 180,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = true,
                RightToLeft = RightToLeft.Yes
            };
            Controls.Add(_txt);

            _rtl = new CheckBox
            {
                Left = 12,
                Top = 220,
                Width = 430,
                // You can set to true if you prefer RLE/PDF by default
                Text = "Use RTL markers (RLE/PDF – usually invisible in SolidWorks)",
                Checked = false
            };
            Controls.Add(_rtl);

            _zwj = new CheckBox
            {
                Left = 12,
                Top = 245,
                Width = 320,
                Text = "Fix disconnected letters (force shaping)",
                Checked = true
            };
            Controls.Add(_zwj);

            var fontLbl = new Label { Left = 12, Top = 275, Width = 50, Text = "Font:" };
            Controls.Add(fontLbl);

            _fontNames = new ComboBox
            {
                Left = 60,
                Top = 270,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            var preferred = new[] { "B Nazanin", "B Zar", "IranNastaliq", "Vazirmatn", "Tahoma", "Segoe UI", "Arial Unicode MS" };
            var installed = FontFamily.Families.Select(f => f.Name).ToList();
            foreach (var f in preferred) if (installed.Contains(f)) _fontNames.Items.Add(f);
            if (_fontNames.Items.Count == 0) foreach (var f in installed) _fontNames.Items.Add(f);
            if (_fontNames.Items.Count > 0) _fontNames.SelectedIndex = 0;
            Controls.Add(_fontNames);

            var sizeLbl = new Label { Left = 270, Top = 275, Width = 40, Text = "Size:" };
            Controls.Add(sizeLbl);

            _fontSize = new NumericUpDown
            {
                Left = 310,
                Top = 272,
                Width = 60,
                Minimum = 6,
                Maximum = 200,
                Increment = 1,
                Value = 12
            };
            Controls.Add(_fontSize);

            _ok = new Button
            {
                Text = "OK",
                Left = 360,
                Top = 305,
                Width = 70,
                DialogResult = DialogResult.OK
            };
            Controls.Add(_ok);

            _cancel = new Button
            {
                Text = "Cancel",
                Left = 438,
                Top = 305,
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

                // Detect marker usage on the *original* string to set the checkbox helpfully.
                bool hadMarkers = ArabicTextUtils.ContainsBidiMarkers(src);
                _rtl.Checked = hadMarkers;

                // Normalize from SolidWorks to logical order for editing.
                _txt.Text = ArabicTextUtils.FromSolidWorks(src);
            }
        }

        // Option: wrap RTL runs with RLE…PDF
        public bool UseRtlMarkers
        {
            get => _rtl.Checked;
            set => _rtl.Checked = value;
        }

        // Option: force Arabic shaping (joins/ligatures)
        public bool InsertJoiners
        {
            get => _zwj.Checked;
            set => _zwj.Checked = value;
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

        // Back-compat for Addin.cs
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
                catch { /* ignore invalid inputs */ }
            }
        }

        private void SelectFont(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                if (_fontNames.Items.Count > 0) _fontNames.SelectedIndex = 0;
                return;
            }
            int idx = -1;
            for (int i = 0; i < _fontNames.Items.Count; i++)
            {
                if (string.Equals(_fontNames.Items[i]?.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    idx = i; break;
                }
            }
            if (idx >= 0) _fontNames.SelectedIndex = idx;
            else if (_fontNames.Items.Count > 0) _fontNames.SelectedIndex = 0;
        }
    }
}
