using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly ComboBox _preset;
        private readonly NumericUpDown _w;
        private readonly NumericUpDown _h;

        private readonly CheckBox _sepMat;
        private readonly CheckBox _onePerMat;
        private readonly CheckBox _filterPreview;

        private readonly RadioButton _rbFast;
        private readonly RadioButton _rbContour1;
        private readonly RadioButton _rbContour2;

        private readonly NumericUpDown _chord;
        private readonly NumericUpDown _snap;

        private readonly Button _ok;
        private readonly Button _cancel;

        private readonly List<SheetPreset> _presets = new List<SheetPreset>
        {
            new SheetPreset("1500 x 3000", 3000, 1500),
            new SheetPreset("1250 x 2500", 2500, 1250),
            new SheetPreset("1000 x 2000", 2000, 1000),
            new SheetPreset("Custom", 3000, 1500),
        };

        public LaserCutRunSettings Settings { get; private set; }

        public LaserCutOptionsForm()
        {
            Text = "Laser nesting options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            // keep OK/Cancel visible
            ClientSize = new Size(680, 420);

            Controls.Add(new Label { Left = 12, Top = 16, Width = 170, Text = "Sheet preset:" });
            _preset = new ComboBox
            {
                Left = 180,
                Top = 12,
                Width = 480,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            foreach (var p in _presets)
                _preset.Items.Add(p.ToString());
            _preset.SelectedIndex = 0;
            _preset.SelectedIndexChanged += (_, __) => ApplyPreset();
            Controls.Add(_preset);

            Controls.Add(new Label { Left = 12, Top = 50, Width = 170, Text = "Width (mm):" });
            _w = new NumericUpDown
            {
                Left = 180,
                Top = 46,
                Width = 140,
                DecimalPlaces = 1,
                Minimum = 100,
                Maximum = 200000,
                Value = 3000
            };
            Controls.Add(_w);

            Controls.Add(new Label { Left = 340, Top = 50, Width = 90, Text = "Height (mm):" });
            _h = new NumericUpDown
            {
                Left = 430,
                Top = 46,
                Width = 140,
                DecimalPlaces = 1,
                Minimum = 100,
                Maximum = 200000,
                Value = 1500
            };
            Controls.Add(_h);

            _sepMat = new CheckBox
            {
                Left = 12,
                Top = 86,
                Width = 650,
                Text = "Separate by EXACT SolidWorks material name",
                Checked = true
            };
            _sepMat.CheckedChanged += (_, __) =>
            {
                bool on = _sepMat.Checked;
                _onePerMat.Enabled = on;
                _filterPreview.Enabled = on;
                if (!on)
                {
                    _onePerMat.Checked = false;
                    _filterPreview.Checked = false;
                }
            };
            Controls.Add(_sepMat);

            _onePerMat = new CheckBox
            {
                Left = 32,
                Top = 112,
                Width = 650,
                Text = "Output one nested DWG per material",
                Checked = true
            };
            Controls.Add(_onePerMat);

            _filterPreview = new CheckBox
            {
                Left = 32,
                Top = 136,
                Width = 650,
                Text = "Keep only that material's source preview (plates + labels) in each output",
                Checked = true
            };
            Controls.Add(_filterPreview);

            var grp = new GroupBox
            {
                Left = 12,
                Top = 170,
                Width = 650,
                Height = 160,
                Text = "Nesting algorithm"
            };
            Controls.Add(grp);

            _rbFast = new RadioButton
            {
                Left = 16,
                Top = 24,
                Width = 620,
                Text = "Fast (Rectangles) — fastest, wastes more sheet",
                Checked = false
            };
            grp.Controls.Add(_rbFast);

            _rbContour1 = new RadioButton
            {
                Left = 16,
                Top = 48,
                Width = 620,
                Text = "Contour (Level 1) — contour + gap offset (good quality, moderate speed)",
                Checked = true
            };
            grp.Controls.Add(_rbContour1);

            _rbContour2 = new RadioButton
            {
                Left = 16,
                Top = 72,
                Width = 620,
                Text = "Contour (Level 2) — NFP/Minkowski touch placement (slowest, best packing)",
                Checked = false
            };
            grp.Controls.Add(_rbContour2);

            grp.Controls.Add(new Label { Left = 36, Top = 110, Width = 160, Text = "Arc chord (mm):" });
            _chord = new NumericUpDown
            {
                Left = 200,
                Top = 106,
                Width = 90,
                DecimalPlaces = 2,
                Minimum = 0.10M,
                Maximum = 5.00M,
                Value = 0.80M
            };
            grp.Controls.Add(_chord);

            grp.Controls.Add(new Label { Left = 320, Top = 110, Width = 160, Text = "Snap tol (mm):" });
            _snap = new NumericUpDown
            {
                Left = 460,
                Top = 106,
                Width = 90,
                DecimalPlaces = 2,
                Minimum = 0.01M,
                Maximum = 0.50M,
                Value = 0.05M
            };
            grp.Controls.Add(_snap);

            var note = new Label
            {
                Left = 12,
                Top = 340,
                Width = 650,
                Height = 24,
                Text = "Note: rotations are always 0/90/180/270. Gap + margin are auto (>= thickness)."
            };
            Controls.Add(note);

            _ok = new Button { Text = "OK", Left = 480, Top = 370, Width = 80, Height = 28 };
            _cancel = new Button { Text = "Cancel", Left = 580, Top = 370, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.Add(_ok);
            Controls.Add(_cancel);

            AcceptButton = _ok;
            CancelButton = _cancel;

            _ok.Click += (_, __) =>
            {
                var chosen = _presets[Math.Max(0, _preset.SelectedIndex)];
                var sheet = new SheetPreset(
                    chosen.Name == "Custom" ? "Custom" : chosen.Name,
                    (double)_w.Value,
                    (double)_h.Value);

                NestingMode mode =
                    _rbContour2.Checked ? NestingMode.ContourLevel2_NFP :
                    _rbContour1.Checked ? NestingMode.ContourLevel1 :
                    NestingMode.FastRectangles;

                Settings = new LaserCutRunSettings
                {
                    DefaultSheet = sheet,

                    SeparateByMaterialExact = _sepMat.Checked,
                    OutputOneDwgPerMaterial = _sepMat.Checked && _onePerMat.Checked,
                    KeepOnlyCurrentMaterialInSourcePreview = _sepMat.Checked && _filterPreview.Checked,

                    Mode = mode,
                    ContourChordMm = (double)_chord.Value,
                    ContourSnapMm = (double)_snap.Value
                };

                DialogResult = DialogResult.OK;
                Close();
            };

            ApplyPreset();
        }

        private void ApplyPreset()
        {
            int idx = Math.Max(0, _preset.SelectedIndex);
            var p = _presets[idx];

            if (!p.Name.Equals("Custom", StringComparison.OrdinalIgnoreCase))
            {
                _w.Value = (decimal)p.WidthMm;
                _h.Value = (decimal)p.HeightMm;
            }
        }
    }

    internal sealed class LaserCutProgressForm : Form
    {
        private readonly ProgressBar _bar;
        private readonly Label _label;

        private readonly int _total;
        private int _done;

        public LaserCutProgressForm(int total)
        {
            _total = Math.Max(1, total);

            Text = "Nesting...";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            ClientSize = new Size(560, 95);

            _label = new Label
            {
                Left = 12,
                Top = 10,
                Width = 536,
                Height = 22,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Starting..."
            };
            Controls.Add(_label);

            _bar = new ProgressBar
            {
                Left = 12,
                Top = 40,
                Width = 536,
                Height = 20,
                Minimum = 0,
                Maximum = _total,
                Value = 0
            };
            Controls.Add(_bar);
        }

        public void Step(string message)
        {
            _done++;
            if (_done > _total) _done = _total;

            if (!string.IsNullOrWhiteSpace(message))
                _label.Text = message;

            _bar.Value = _done;

            _label.Refresh();
            _bar.Refresh();
            Application.DoEvents();
        }

        // Optional compatibility (if any of your other code calls SetStatus)
        public void SetStatus(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
                _label.Text = message;

            _label.Refresh();
            Application.DoEvents();
        }
    }
}
