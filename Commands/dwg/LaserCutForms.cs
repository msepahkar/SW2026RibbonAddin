using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly string _folder;
        private readonly List<LaserNestJob> _jobs; // sorted

        private readonly DataGridView _grid;

        private readonly RadioButton _rbFast;
        private readonly RadioButton _rbContour1;
        private readonly RadioButton _rbContour2;

        private readonly NumericUpDown _chord;
        private readonly NumericUpDown _snap;

        private readonly Button _btnAll;
        private readonly Button _btnNone;

        private readonly Button _ok;
        private readonly Button _cancel;

        private readonly List<SheetPreset> _presets = new List<SheetPreset>
        {
            new SheetPreset("1500 x 3000 mm", 3000, 1500),
            new SheetPreset("1250 x 2500 mm", 2500, 1250),
            new SheetPreset("1000 x 2000 mm", 2000, 1000),
            new SheetPreset("Custom", 0, 0),
        };

        public LaserCutRunSettings Settings { get; private set; }

        public List<LaserNestJob> SelectedJobs { get; private set; }

        public LaserCutOptionsForm(string folder, List<LaserNestJob> jobs)
        {
            _folder = folder ?? "";
            _jobs = (jobs ?? new List<LaserNestJob>())
                .OrderBy(j => j.MaterialExact ?? "", StringComparer.Ordinal)
                .ThenBy(j => j.ThicknessMm <= 0 ? double.MaxValue : j.ThicknessMm)
                .ThenBy(j => j.ThicknessFileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            Text = "Laser nesting options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            ClientSize = new Size(980, 620);

            var title = new Label
            {
                Left = 12,
                Top = 10,
                Width = 950,
                Height = 22,
                Text = "Select which (Material × Thickness) runs to nest, and set sheet size per item:"
            };
            Controls.Add(title);

            _grid = new DataGridView
            {
                Left = 12,
                Top = 38,
                Width = 950,
                Height = 360,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(_grid);

            BuildGridColumns();
            PopulateGrid();

            _grid.CurrentCellDirtyStateChanged += (_, __) =>
            {
                if (_grid.IsCurrentCellDirty)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CellEndEdit += Grid_CellEndEdit;
            _grid.DataError += (_, __) => { /* ignore combo parse errors */ };

            _btnAll = new Button { Left = 12, Top = 406, Width = 120, Height = 28, Text = "Select All" };
            _btnNone = new Button { Left = 140, Top = 406, Width = 120, Height = 28, Text = "Select None" };
            Controls.Add(_btnAll);
            Controls.Add(_btnNone);

            _btnAll.Click += (_, __) => SetAllEnabled(true);
            _btnNone.Click += (_, __) => SetAllEnabled(false);

            var grp = new GroupBox
            {
                Left = 12,
                Top = 445,
                Width = 950,
                Height = 120,
                Text = "Nesting algorithm"
            };
            Controls.Add(grp);

            _rbFast = new RadioButton
            {
                Left = 16,
                Top = 24,
                Width = 900,
                Text = "Fast (Rectangles) — fastest, wastes more sheet",
                Checked = false
            };
            grp.Controls.Add(_rbFast);

            _rbContour1 = new RadioButton
            {
                Left = 16,
                Top = 48,
                Width = 900,
                Text = "Contour (Level 1) — contour + gap offset (good quality, moderate speed)",
                Checked = true
            };
            grp.Controls.Add(_rbContour1);

            _rbContour2 = new RadioButton
            {
                Left = 16,
                Top = 72,
                Width = 900,
                Text = "Contour (Level 2) — NFP/Minkowski touch placement (slowest, best packing)",
                Checked = false
            };
            grp.Controls.Add(_rbContour2);

            grp.Controls.Add(new Label { Left = 36, Top = 96, Width = 160, Text = "Arc chord (mm):" });
            _chord = new NumericUpDown
            {
                Left = 200,
                Top = 92,
                Width = 90,
                DecimalPlaces = 2,
                Minimum = 0.10M,
                Maximum = 5.00M,
                Value = 0.80M
            };
            grp.Controls.Add(_chord);

            grp.Controls.Add(new Label { Left = 310, Top = 96, Width = 110, Text = "Snap tol (mm):" });
            _snap = new NumericUpDown
            {
                Left = 430,
                Top = 92,
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
                Top = 570,
                Width = 950,
                Height = 22,
                Text = "Note: rotations are always 0/90/180/270. Gap+margin are auto (>= thickness)."
            };
            Controls.Add(note);

            _ok = new Button { Text = "OK", Left = 780, Top = 592, Width = 80, Height = 28 };
            _cancel = new Button { Text = "Cancel", Left = 882, Top = 592, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.Add(_ok);
            Controls.Add(_cancel);

            AcceptButton = _ok;
            CancelButton = _cancel;

            _ok.Click += (_, __) => OnOk();

            // Ensure sheet values are loaded from last memory before user sees UI
            ApplyRememberedSheetsIntoGrid();
        }

        private void BuildGridColumns()
        {
            _grid.Columns.Clear();

            var colOn = new DataGridViewCheckBoxColumn
            {
                Name = "Enabled",
                HeaderText = "",
                Width = 36
            };
            _grid.Columns.Add(colOn);

            var colMat = new DataGridViewTextBoxColumn
            {
                Name = "Material",
                HeaderText = "Material (EXACT from SolidWorks)",
                Width = 320,
                ReadOnly = true
            };
            _grid.Columns.Add(colMat);

            var colThk = new DataGridViewTextBoxColumn
            {
                Name = "Thickness",
                HeaderText = "Thk (mm)",
                Width = 80,
                ReadOnly = true
            };
            _grid.Columns.Add(colThk);

            var colFile = new DataGridViewTextBoxColumn
            {
                Name = "File",
                HeaderText = "Source DWG",
                Width = 190,
                ReadOnly = true
            };
            _grid.Columns.Add(colFile);

            var colPreset = new DataGridViewComboBoxColumn
            {
                Name = "Preset",
                HeaderText = "Sheet preset",
                Width = 160,
                FlatStyle = FlatStyle.Flat
            };
            foreach (var p in _presets)
                colPreset.Items.Add(p.Name);
            _grid.Columns.Add(colPreset);

            var colW = new DataGridViewTextBoxColumn
            {
                Name = "W",
                HeaderText = "W (mm)",
                Width = 80
            };
            _grid.Columns.Add(colW);

            var colH = new DataGridViewTextBoxColumn
            {
                Name = "H",
                HeaderText = "H (mm)",
                Width = 80
            };
            _grid.Columns.Add(colH);
        }

        private void PopulateGrid()
        {
            _grid.Rows.Clear();

            foreach (var j in _jobs)
            {
                var rowIndex = _grid.Rows.Add();
                var row = _grid.Rows[rowIndex];

                row.Tag = j;

                row.Cells["Enabled"].Value = true;
                row.Cells["Material"].Value = j.MaterialExact ?? "UNKNOWN";
                row.Cells["Thickness"].Value = j.ThicknessMm > 0 ? j.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture) : "?";
                row.Cells["File"].Value = j.ThicknessFileName;

                // placeholders; real values loaded later by ApplyRememberedSheetsIntoGrid
                row.Cells["Preset"].Value = _presets[0].Name;
                row.Cells["W"].Value = _presets[0].WidthMm.ToString("0.###", CultureInfo.InvariantCulture);
                row.Cells["H"].Value = _presets[0].HeightMm.ToString("0.###", CultureInfo.InvariantCulture);
            }
        }

        private void ApplyRememberedSheetsIntoGrid()
        {
            // global default
            var global = LaserCutUiMemory.LoadGlobalDefaultSheet(_presets[0]);

            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (!(row.Tag is LaserNestJob job))
                    continue;

                var remembered = LaserCutUiMemory.LoadSheetFor(job.MaterialExact, job.ThicknessMm, global);

                int presetIdx = FindPresetIndex(remembered.WidthMm, remembered.HeightMm);
                string presetName = presetIdx >= 0 ? _presets[presetIdx].Name : "Custom";

                row.Cells["Preset"].Value = presetName;
                row.Cells["W"].Value = remembered.WidthMm.ToString("0.###", CultureInfo.InvariantCulture);
                row.Cells["H"].Value = remembered.HeightMm.ToString("0.###", CultureInfo.InvariantCulture);
            }
        }

        private int FindPresetIndex(double w, double h)
        {
            const double eps = 0.001;
            for (int i = 0; i < _presets.Count; i++)
            {
                var p = _presets[i];
                if (string.Equals(p.Name, "Custom", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (Math.Abs(p.WidthMm - w) < eps && Math.Abs(p.HeightMm - h) < eps)
                    return i;
            }
            return -1;
        }

        private void SetAllEnabled(bool enabled)
        {
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells["Enabled"].Value = enabled;
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var row = _grid.Rows[e.RowIndex];
            string colName = _grid.Columns[e.ColumnIndex].Name;

            if (colName == "Preset")
            {
                string presetName = (row.Cells["Preset"].Value as string) ?? "";
                var preset = _presets.FirstOrDefault(p => string.Equals(p.Name, presetName, StringComparison.OrdinalIgnoreCase));

                if (preset.Name != null && !string.Equals(preset.Name, "Custom", StringComparison.OrdinalIgnoreCase))
                {
                    row.Cells["W"].Value = preset.WidthMm.ToString("0.###", CultureInfo.InvariantCulture);
                    row.Cells["H"].Value = preset.HeightMm.ToString("0.###", CultureInfo.InvariantCulture);
                }
            }
        }

        private void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var row = _grid.Rows[e.RowIndex];
            string colName = _grid.Columns[e.ColumnIndex].Name;

            if (colName == "W" || colName == "H")
            {
                // validate number and auto-switch preset to Custom if mismatch
                if (!TryParseCellDouble(row.Cells["W"].Value, out double w) || w <= 0 ||
                    !TryParseCellDouble(row.Cells["H"].Value, out double h) || h <= 0)
                {
                    MessageBox.Show("Width/Height must be valid positive numbers.", "Invalid sheet size",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    // reset to a safe preset
                    row.Cells["Preset"].Value = _presets[0].Name;
                    row.Cells["W"].Value = _presets[0].WidthMm.ToString("0.###", CultureInfo.InvariantCulture);
                    row.Cells["H"].Value = _presets[0].HeightMm.ToString("0.###", CultureInfo.InvariantCulture);
                    return;
                }

                int presetIdx = FindPresetIndex(w, h);
                if (presetIdx < 0)
                    row.Cells["Preset"].Value = "Custom";
                else
                    row.Cells["Preset"].Value = _presets[presetIdx].Name;
            }
        }

        private static bool TryParseCellDouble(object v, out double value)
        {
            value = 0.0;
            if (v == null)
                return false;

            string s = v.ToString().Trim();
            if (s.Length == 0)
                return false;

            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private void OnOk()
        {
            var selected = new List<LaserNestJob>();

            // Build settings (3 checkboxes enforced = true)
            NestingMode mode =
                _rbContour2.Checked ? NestingMode.ContourLevel2_NFP :
                _rbContour1.Checked ? NestingMode.ContourLevel1 :
                NestingMode.FastRectangles;

            var settings = new LaserCutRunSettings
            {
                SeparateByMaterialExact = true,
                OutputOneDwgPerMaterial = true,
                KeepOnlyCurrentMaterialInSourcePreview = true,

                Mode = mode,
                ContourChordMm = (double)_chord.Value,
                ContourSnapMm = (double)_snap.Value,

                DefaultSheet = _presets[0] // not very important now
            };

            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (!(row.Tag is LaserNestJob job))
                    continue;

                bool enabled = row.Cells["Enabled"].Value is bool b && b;
                job.Enabled = enabled;

                if (!TryParseCellDouble(row.Cells["W"].Value, out double w) || w <= 0 ||
                    !TryParseCellDouble(row.Cells["H"].Value, out double h) || h <= 0)
                {
                    MessageBox.Show("One or more sheet sizes are invalid. Fix them before pressing OK.",
                        "Invalid sheet size",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                string presetName = (row.Cells["Preset"].Value as string) ?? "Custom";
                job.Sheet = new SheetPreset(presetName, w, h);

                // Remember per job even if disabled (so user doesn't lose their editing)
                LaserCutUiMemory.SaveSheetFor(job.MaterialExact, job.ThicknessMm, job.Sheet);

                if (enabled)
                    selected.Add(job);
            }

            if (selected.Count == 0)
            {
                MessageBox.Show("Nothing selected. Check at least one item.", "Laser nesting",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Save global default as the first enabled item (simple, predictable)
            LaserCutUiMemory.SaveGlobalDefaultSheet(selected[0].Sheet);

            Settings = settings;
            SelectedJobs = selected;

            DialogResult = DialogResult.OK;
            Close();
        }
    }


namespace SW2026RibbonAddin.Commands
    {
        internal sealed class LaserCutProgressForm : Form
        {
            private readonly Label _lblHeader;
            private readonly Label _lblTask;
            private readonly Label _lblCounts;
            private readonly Label _lblStatus;
            private readonly ProgressBar _bar;
            private readonly Button _btnCancel;

            private volatile bool _cancelRequested;

            private int _batchTotal;
            private int _batchIndex;

            private int _totalParts;
            private int _placedParts;
            private int _sheetsUsed;

            private string _file;
            private string _material;
            private double _thickness;
            private NestingMode _mode;
            private double _sheetW;
            private double _sheetH;

            public bool IsCancellationRequested => _cancelRequested;

            public LaserCutProgressForm()
            {
                Text = "Nesting...";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;
                Width = 520;
                Height = 190;

                _lblHeader = new Label { Left = 12, Top = 10, Width = 480, Height = 18, Text = "Nesting..." };
                _lblTask = new Label { Left = 12, Top = 32, Width = 480, Height = 36, Text = "" };
                _lblCounts = new Label { Left = 12, Top = 70, Width = 480, Height = 18, Text = "" };

                _bar = new ProgressBar { Left = 12, Top = 92, Width = 480, Height = 18, Minimum = 0, Maximum = 100, Value = 0 };

                _lblStatus = new Label { Left = 12, Top = 114, Width = 480, Height = 18, Text = "" };

                _btnCancel = new Button { Left = 402, Top = 136, Width = 90, Height = 26, Text = "Cancel" };
                _btnCancel.Click += (s, e) => RequestCancel();

                Controls.Add(_lblHeader);
                Controls.Add(_lblTask);
                Controls.Add(_lblCounts);
                Controls.Add(_bar);
                Controls.Add(_lblStatus);
                Controls.Add(_btnCancel);

                // If user clicks [X], treat as cancel request (don’t kill the process abruptly)
                FormClosing += (s, e) =>
                {
                    if (!_cancelRequested)
                    {
                        _cancelRequested = true;
                        _btnCancel.Enabled = false;
                        _lblStatus.Text = "Cancelling...";
                        PumpUI();
                    }
                    // allow closing
                };
            }

            public void BeginBatch(int totalTasks)
            {
                UI(() =>
                {
                    _batchTotal = Math.Max(1, totalTasks);
                    _batchIndex = 0;

                    _lblHeader.Text = "Nesting batch started";
                    _lblStatus.Text = "";
                    _bar.Value = 0;
                    _btnCancel.Enabled = true;
                });
            }

            public void BeginTask(
                int taskIndex,
                int totalTasks,
                string thicknessFileName,
                string materialExact,
                double thicknessMm,
                int totalParts,
                NestingMode mode,
                double sheetWmm,
                double sheetHmm)
            {
                UI(() =>
                {
                    _batchIndex = Math.Max(1, taskIndex);
                    _batchTotal = Math.Max(1, totalTasks);

                    _file = thicknessFileName ?? "";
                    _material = materialExact ?? "UNKNOWN";
                    _thickness = thicknessMm;
                    _mode = mode;
                    _sheetW = sheetWmm;
                    _sheetH = sheetHmm;

                    _totalParts = Math.Max(0, totalParts);
                    _placedParts = 0;
                    _sheetsUsed = 1;

                    _lblHeader.Text = $"Nesting...  Task {_batchIndex}/{_batchTotal}";
                    _lblTask.Text =
                        $"{_file}\r\n" +
                        $"{_material} | {(_thickness > 0 ? _thickness.ToString("0.###") : "?")} mm | {_mode} | Sheet {_sheetW:0.###}×{_sheetH:0.###}";

                    _lblCounts.Text = $"Placed {_placedParts}/{_totalParts}   Sheets: {_sheetsUsed}";
                    _lblStatus.Text = "";

                    _bar.Minimum = 0;
                    _bar.Maximum = Math.Max(1, _totalParts);
                    _bar.Value = 0;

                    _btnCancel.Enabled = true;
                });
            }

            public void ReportPlaced(int placed, int total, int sheetsUsed)
            {
                UI(() =>
                {
                    _placedParts = Math.Max(0, placed);
                    _totalParts = Math.Max(0, total);
                    _sheetsUsed = Math.Max(1, sheetsUsed);

                    if (_bar.Maximum != Math.Max(1, _totalParts))
                        _bar.Maximum = Math.Max(1, _totalParts);

                    _bar.Value = Math.Min(_bar.Maximum, Math.Max(_bar.Minimum, _placedParts));
                    _lblCounts.Text = $"Placed {_placedParts}/{_totalParts}   Sheets: {_sheetsUsed}";
                });

                ThrowIfCancelled();
            }

            public void EndTask(int doneTasks)
            {
                UI(() =>
                {
                    _lblStatus.Text = $"Finished task {doneTasks}/{_batchTotal}";
                });

                ThrowIfCancelled();
            }

            public void SetStatus(string message)
            {
                UI(() =>
                {
                    _lblStatus.Text = message ?? "";
                });

                ThrowIfCancelled();
            }

            public void ThrowIfCancelled()
            {
                if (_cancelRequested)
                    throw new OperationCanceledException("User cancelled nesting.");
            }

            private void RequestCancel()
            {
                _cancelRequested = true;
                UI(() =>
                {
                    _btnCancel.Enabled = false;
                    _lblStatus.Text = "Cancelling...";
                });
            }

            private void UI(Action action)
            {
                if (IsDisposed) return;

                if (InvokeRequired)
                {
                    try { BeginInvoke(action); } catch { }
                    return;
                }

                action();
                PumpUI();
            }

            private void PumpUI()
            {
                // IMPORTANT: keeps the form responsive when nesting runs on the same thread
                try { System.Windows.Forms.Application.DoEvents(); } catch { }
            }
        }
    }

    internal static class LaserCutUiMemory
    {
        private const string BaseKey = @"Software\SW2026RibbonAddin\LaserNesting";

        public static SheetPreset LoadGlobalDefaultSheet(SheetPreset fallback)
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(BaseKey + @"\Global"))
                {
                    if (k == null) return fallback;

                    string name = (k.GetValue("Preset", fallback.Name) as string) ?? fallback.Name;
                    double w = ReadDouble(k, "W", fallback.WidthMm);
                    double h = ReadDouble(k, "H", fallback.HeightMm);

                    if (w > 0 && h > 0)
                        return new SheetPreset(name, w, h);
                }
            }
            catch { }

            return fallback;
        }

        public static void SaveGlobalDefaultSheet(SheetPreset sheet)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(BaseKey + @"\Global"))
                {
                    if (k == null) return;

                    k.SetValue("Preset", sheet.Name ?? "Custom");
                    k.SetValue("W", sheet.WidthMm.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("H", sheet.HeightMm.ToString("R", CultureInfo.InvariantCulture));
                }
            }
            catch { }
        }

        public static SheetPreset LoadSheetFor(string materialExact, double thicknessMm, SheetPreset fallback)
        {
            try
            {
                string key = JobKey(materialExact, thicknessMm);
                using (var k = Registry.CurrentUser.OpenSubKey(BaseKey + @"\Jobs\" + key))
                {
                    if (k == null) return fallback;

                    string name = (k.GetValue("Preset", fallback.Name) as string) ?? fallback.Name;
                    double w = ReadDouble(k, "W", fallback.WidthMm);
                    double h = ReadDouble(k, "H", fallback.HeightMm);

                    if (w > 0 && h > 0)
                        return new SheetPreset(name, w, h);
                }
            }
            catch { }

            return fallback;
        }

        public static void SaveSheetFor(string materialExact, double thicknessMm, SheetPreset sheet)
        {
            try
            {
                string key = JobKey(materialExact, thicknessMm);
                using (var k = Registry.CurrentUser.CreateSubKey(BaseKey + @"\Jobs\" + key))
                {
                    if (k == null) return;

                    k.SetValue("Material", materialExact ?? "");
                    k.SetValue("Thickness", thicknessMm.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("Preset", sheet.Name ?? "Custom");
                    k.SetValue("W", sheet.WidthMm.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("H", sheet.HeightMm.ToString("R", CultureInfo.InvariantCulture));
                }
            }
            catch { }
        }

        private static double ReadDouble(RegistryKey k, string name, double fallback)
        {
            try
            {
                var v = k.GetValue(name);
                if (v == null) return fallback;

                string s = v.ToString();
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                    return d;
            }
            catch { }

            return fallback;
        }

        private static string JobKey(string materialExact, double thicknessMm)
        {
            string input = (materialExact ?? "") + "|" + thicknessMm.ToString("0.###", CultureInfo.InvariantCulture);
            using (var md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));
                var sb = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                    sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }
    }
}
