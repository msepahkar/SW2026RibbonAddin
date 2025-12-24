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

        private readonly SplitContainer _split;
        private readonly TreeView _tree;

        // Details panel (per selected job)
        private readonly Label _lblMaterial;
        private readonly Label _lblThickness;
        private readonly Label _lblSource;

        private readonly ComboBox _cbPreset;
        private readonly NumericUpDown _numW;
        private readonly NumericUpDown _numH;

        private readonly Button _btnApplyThickness;
        private readonly Button _btnApplyAllSheets;

        private readonly Button _btnAll;
        private readonly Button _btnNone;

        private readonly RadioButton _rbFast;
        private readonly RadioButton _rbContour1;
        private readonly RadioButton _rbContour2;

        private readonly NumericUpDown _chord;
        private readonly NumericUpDown _snap;

        private readonly Button _ok;
        private readonly Button _cancel;

        private readonly Dictionary<LaserNestJob, TreeNode> _nodeByJob = new Dictionary<LaserNestJob, TreeNode>();

        private bool _suppressTreeEvents;
        private bool _suppressDetailEvents;

        private LaserNestJob _selectedJob;

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
                .OrderBy(j => j.ThicknessMm <= 0 ? double.MaxValue : j.ThicknessMm)
                .ThenBy(j => j.ThicknessFileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ThenBy(j => j.MaterialExact ?? "", StringComparer.Ordinal)
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
                Text = "Select which (Thickness × Material) runs to nest, and set sheet size per item:"
            };
            Controls.Add(title);

            // Apply remembered sheets BEFORE building UI nodes
            ApplyRememberedSheetsIntoJobs();

            _split = new SplitContainer
            {
                Left = 12,
                Top = 38,
                Width = 950,
                Height = 360,
                Orientation = Orientation.Vertical,
                SplitterDistance = 480,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(_split);

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                CheckBoxes = true,
                HideSelection = false
            };
            _tree.AfterCheck += Tree_AfterCheck;
            _tree.AfterSelect += Tree_AfterSelect;
            _split.Panel1.Controls.Add(_tree);

            // Details panel
            var details = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Selected item"
            };
            _split.Panel2.Controls.Add(details);

            _lblMaterial = new Label { Left = 12, Top = 24, Width = 430, Height = 18, Text = "Material: -" };
            _lblThickness = new Label { Left = 12, Top = 44, Width = 430, Height = 18, Text = "Thickness: -" };
            _lblSource = new Label { Left = 12, Top = 64, Width = 430, Height = 18, Text = "Source DWG: -" };

            details.Controls.Add(_lblMaterial);
            details.Controls.Add(_lblThickness);
            details.Controls.Add(_lblSource);

            details.Controls.Add(new Label { Left = 12, Top = 96, Width = 110, Height = 18, Text = "Sheet preset:" });

            _cbPreset = new ComboBox
            {
                Left = 128,
                Top = 92,
                Width = 220,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            foreach (var p in _presets)
                _cbPreset.Items.Add(p.Name);
            details.Controls.Add(_cbPreset);

            details.Controls.Add(new Label { Left = 12, Top = 128, Width = 110, Height = 18, Text = "Width (mm):" });
            _numW = new NumericUpDown
            {
                Left = 128,
                Top = 124,
                Width = 120,
                DecimalPlaces = 3,
                Minimum = 1,
                Maximum = 200000,
                Value = 3000,
                Increment = 10
            };
            details.Controls.Add(_numW);

            details.Controls.Add(new Label { Left = 260, Top = 128, Width = 90, Height = 18, Text = "Height (mm):" });
            _numH = new NumericUpDown
            {
                Left = 340,
                Top = 124,
                Width = 120,
                DecimalPlaces = 3,
                Minimum = 1,
                Maximum = 200000,
                Value = 1500,
                Increment = 10
            };
            details.Controls.Add(_numH);

            _btnApplyThickness = new Button { Left = 12, Top = 160, Width = 220, Height = 28, Text = "Apply sheet to this thickness" };
            _btnApplyAllSheets = new Button { Left = 240, Top = 160, Width = 220, Height = 28, Text = "Apply sheet to ALL jobs" };
            details.Controls.Add(_btnApplyThickness);
            details.Controls.Add(_btnApplyAllSheets);

            _cbPreset.SelectedIndexChanged += (_, __) => OnPresetChanged();
            _numW.ValueChanged += (_, __) => OnSheetDimChanged();
            _numH.ValueChanged += (_, __) => OnSheetDimChanged();

            _btnApplyThickness.Click += (_, __) => ApplySelectedSheetToThickness();
            _btnApplyAllSheets.Click += (_, __) => ApplySelectedSheetToAll();

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

            BuildTree();

            // Preselect the first job so the user sees sheet controls immediately
            if (_tree.Nodes.Count > 0)
            {
                var first = _tree.Nodes[0];
                if (first.Nodes.Count > 0)
                    _tree.SelectedNode = first.Nodes[0];
                else
                    _tree.SelectedNode = first;
            }
            else
            {
                SetDetailsEnabled(false);
            }
        }

        private void ApplyRememberedSheetsIntoJobs()
        {
            var global = LaserCutUiMemory.LoadGlobalDefaultSheet(_presets[0]);

            foreach (var job in _jobs)
            {
                var remembered = LaserCutUiMemory.LoadSheetFor(job.MaterialExact, job.ThicknessMm, global);

                // Normalize name against our preset list (so the combo always has a valid selection)
                int presetIdx = FindPresetIndex(remembered.WidthMm, remembered.HeightMm);
                string presetName = presetIdx >= 0 ? _presets[presetIdx].Name : "Custom";

                double w = remembered.WidthMm > 0 ? remembered.WidthMm : global.WidthMm;
                double h = remembered.HeightMm > 0 ? remembered.HeightMm : global.HeightMm;

                if (w <= 0) w = _presets[0].WidthMm;
                if (h <= 0) h = _presets[0].HeightMm;

                job.Sheet = new SheetPreset(presetName, w, h);
            }
        }

        private void BuildTree()
        {
            _tree.BeginUpdate();
            _tree.Nodes.Clear();
            _nodeByJob.Clear();

            var groups = _jobs
                .GroupBy(j => (j.ThicknessFileName ?? ""))
                .Select(g => new
                {
                    FileName = g.Key,
                    Thickness = g.FirstOrDefault()?.ThicknessMm ?? 0.0,
                    Jobs = g.OrderBy(j => j.MaterialExact ?? "", StringComparer.Ordinal).ToList()
                })
                .OrderBy(g => g.Thickness <= 0 ? double.MaxValue : g.Thickness)
                .ThenBy(g => g.FileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var g in groups)
            {
                int total = g.Jobs.Count;
                int enabled = g.Jobs.Count(j => j != null && j.Enabled);

                var root = new TreeNode(BuildRootText(g.FileName, g.Thickness, enabled, total))
                {
                    Tag = null,
                    Checked = total > 0 && enabled == total
                };

                foreach (var job in g.Jobs)
                {
                    if (job == null) continue;

                    var node = new TreeNode(BuildJobText(job))
                    {
                        Tag = job,
                        Checked = job.Enabled
                    };

                    root.Nodes.Add(node);
                    _nodeByJob[job] = node;
                }

                root.Expand();
                _tree.Nodes.Add(root);
            }

            _tree.EndUpdate();
        }

        private static string BuildRootText(string thicknessFileName, double thicknessMm, int enabled, int total)
        {
            string thk = thicknessMm > 0 ? thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) : "?";
            string file = string.IsNullOrWhiteSpace(thicknessFileName) ? "(no file)" : thicknessFileName.Trim();
            return $"{file}  ({thk} mm)   [{enabled}/{total}]";
        }

        private string BuildJobText(LaserNestJob job)
        {
            if (job == null)
                return "(null)";

            string mat = string.IsNullOrWhiteSpace(job.MaterialExact) ? "UNKNOWN" : job.MaterialExact.Trim();

            double w = job.Sheet.WidthMm;
            double h = job.Sheet.HeightMm;

            string presetName = job.Sheet.Name ?? "";
            if (string.IsNullOrWhiteSpace(presetName))
            {
                int idx = FindPresetIndex(w, h);
                presetName = idx >= 0 ? _presets[idx].Name : "Custom";
            }

            string dims = $"{w:0.###}×{h:0.###} mm";
            if (!string.Equals(presetName, "Custom", StringComparison.OrdinalIgnoreCase))
                return $"{mat}   —   {dims}   ({presetName})";

            return $"{mat}   —   {dims}";
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

        private void Tree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_suppressTreeEvents)
                return;

            _suppressTreeEvents = true;

            try
            {
                if (e.Node?.Tag is LaserNestJob job)
                {
                    job.Enabled = e.Node.Checked;

                    UpdateJobNode(job);

                    if (e.Node.Parent != null)
                        UpdateRootNode(e.Node.Parent);
                }
                else
                {
                    // Root node toggled => apply to children
                    if (e.Node != null)
                    {
                        foreach (TreeNode child in e.Node.Nodes)
                        {
                            child.Checked = e.Node.Checked;

                            if (child.Tag is LaserNestJob cj)
                            {
                                cj.Enabled = child.Checked;
                                UpdateJobNode(cj);
                            }
                        }

                        UpdateRootNode(e.Node);
                    }
                }
            }
            finally
            {
                _suppressTreeEvents = false;
            }
        }

        private void Tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node?.Tag is LaserNestJob job)
            {
                LoadJobIntoDetails(job);
            }
            else
            {
                _selectedJob = null;
                _lblMaterial.Text = "Material: -";
                _lblThickness.Text = "Thickness: -";
                _lblSource.Text = "Source DWG: -";
                SetDetailsEnabled(false);
            }
        }

        private void LoadJobIntoDetails(LaserNestJob job)
        {
            _selectedJob = job;
            if (_selectedJob == null)
            {
                SetDetailsEnabled(false);
                return;
            }

            SetDetailsEnabled(true);

            _suppressDetailEvents = true;
            try
            {
                _lblMaterial.Text = "Material: " + (string.IsNullOrWhiteSpace(job.MaterialExact) ? "UNKNOWN" : job.MaterialExact);
                _lblThickness.Text = "Thickness: " + (job.ThicknessMm > 0 ? job.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture) : "?") + " mm";
                _lblSource.Text = "Source DWG: " + (job.ThicknessFileName ?? "");

                double w = job.Sheet.WidthMm > 0 ? job.Sheet.WidthMm : _presets[0].WidthMm;
                double h = job.Sheet.HeightMm > 0 ? job.Sheet.HeightMm : _presets[0].HeightMm;

                // Clamp to numeric control range
                w = Math.Min(w, (double)_numW.Maximum);
                h = Math.Min(h, (double)_numH.Maximum);

                _numW.Value = (decimal)w;
                _numH.Value = (decimal)h;

                int presetIdx = FindPresetIndex(w, h);
                string presetName = presetIdx >= 0 ? _presets[presetIdx].Name : "Custom";

                if (_cbPreset.Items.Contains(presetName))
                    _cbPreset.SelectedItem = presetName;
                else
                    _cbPreset.SelectedItem = "Custom";
            }
            finally
            {
                _suppressDetailEvents = false;
            }
        }

        private void SetDetailsEnabled(bool enabled)
        {
            _cbPreset.Enabled = enabled;
            _numW.Enabled = enabled;
            _numH.Enabled = enabled;
            _btnApplyThickness.Enabled = enabled;
            _btnApplyAllSheets.Enabled = enabled;
        }

        private void OnPresetChanged()
        {
            if (_suppressDetailEvents)
                return;

            if (_selectedJob == null)
                return;

            string presetName = (_cbPreset.SelectedItem as string) ?? "Custom";
            var preset = _presets.FirstOrDefault(p => string.Equals(p.Name, presetName, StringComparison.OrdinalIgnoreCase));

            if (!string.Equals(preset.Name, "Custom", StringComparison.OrdinalIgnoreCase) &&
                preset.WidthMm > 0 && preset.HeightMm > 0)
            {
                _suppressDetailEvents = true;
                try
                {
                    _numW.Value = (decimal)Math.Min(preset.WidthMm, (double)_numW.Maximum);
                    _numH.Value = (decimal)Math.Min(preset.HeightMm, (double)_numH.Maximum);
                }
                finally
                {
                    _suppressDetailEvents = false;
                }

                SetSelectedJobSheet(preset.Name, preset.WidthMm, preset.HeightMm);
                return;
            }

            // Custom preset: keep numeric values
            SetSelectedJobSheet("Custom", (double)_numW.Value, (double)_numH.Value);
        }

        private void OnSheetDimChanged()
        {
            if (_suppressDetailEvents)
                return;

            if (_selectedJob == null)
                return;

            double w = (double)_numW.Value;
            double h = (double)_numH.Value;

            int presetIdx = FindPresetIndex(w, h);
            string presetName = presetIdx >= 0 ? _presets[presetIdx].Name : "Custom";

            _suppressDetailEvents = true;
            try
            {
                if (_cbPreset.Items.Contains(presetName))
                    _cbPreset.SelectedItem = presetName;
                else
                    _cbPreset.SelectedItem = "Custom";
            }
            finally
            {
                _suppressDetailEvents = false;
            }

            SetSelectedJobSheet(presetName, w, h);
        }

        private void SetSelectedJobSheet(string presetName, double w, double h)
        {
            if (_selectedJob == null)
                return;

            w = Math.Max(1.0, w);
            h = Math.Max(1.0, h);

            _selectedJob.Sheet = new SheetPreset(presetName ?? "Custom", w, h);

            UpdateJobNode(_selectedJob);
        }

        private void ApplySelectedSheetToThickness()
        {
            if (_selectedJob == null)
                return;

            string file = _selectedJob.ThicknessFileName ?? "";

            foreach (var job in _jobs.Where(j => j != null && string.Equals(j.ThicknessFileName ?? "", file, StringComparison.OrdinalIgnoreCase)))
            {
                job.Sheet = _selectedJob.Sheet;
                UpdateJobNode(job);
            }
        }

        private void ApplySelectedSheetToAll()
        {
            if (_selectedJob == null)
                return;

            foreach (var job in _jobs.Where(j => j != null))
            {
                job.Sheet = _selectedJob.Sheet;
                UpdateJobNode(job);
            }
        }

        private void UpdateJobNode(LaserNestJob job)
        {
            if (job == null) return;

            if (_nodeByJob.TryGetValue(job, out var node) && node != null)
                node.Text = BuildJobText(job);
        }

        private void UpdateRootNode(TreeNode root)
        {
            if (root == null) return;
            if (root.Nodes == null) return;

            int total = root.Nodes.Count;
            int enabled = 0;
            double thickness = 0.0;
            string file = root.Text;

            foreach (TreeNode child in root.Nodes)
            {
                if (child.Checked) enabled++;

                if (thickness <= 0 && child.Tag is LaserNestJob j)
                    thickness = j.ThicknessMm;

                if (child.Tag is LaserNestJob j2 && !string.IsNullOrWhiteSpace(j2.ThicknessFileName))
                    file = j2.ThicknessFileName;
            }

            // Root checkbox = "all selected"
            root.Checked = total > 0 && enabled == total;
            root.Text = BuildRootText(file, thickness, enabled, total);
        }

        private void SetAllEnabled(bool enabled)
        {
            _suppressTreeEvents = true;

            try
            {
                _tree.BeginUpdate();

                foreach (TreeNode root in _tree.Nodes)
                {
                    foreach (TreeNode child in root.Nodes)
                    {
                        child.Checked = enabled;

                        if (child.Tag is LaserNestJob job)
                            job.Enabled = enabled;
                    }

                    UpdateRootNode(root);
                }
            }
            finally
            {
                _tree.EndUpdate();
                _suppressTreeEvents = false;
            }
        }

        private void OnOk()
        {
            var selected = _jobs.Where(j => j != null && j.Enabled).ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show("Nothing selected. Check at least one item.", "Laser nesting",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Validate sheet sizes (even for disabled jobs, to preserve memory safely)
            foreach (var job in _jobs)
            {
                if (job == null) continue;

                if (job.Sheet.WidthMm <= 0 || job.Sheet.HeightMm <= 0)
                {
                    MessageBox.Show("One or more sheet sizes are invalid. Fix them before pressing OK.",
                        "Invalid sheet size",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
            }

            // Save global default as the first enabled item (simple, predictable)
            LaserCutUiMemory.SaveGlobalDefaultSheet(selected[0].Sheet);

            // Save per-job sheet dims (even if disabled, so user doesn't lose editing)
            foreach (var job in _jobs)
            {
                if (job == null) continue;
                LaserCutUiMemory.SaveSheetFor(job.MaterialExact, job.ThicknessMm, job.Sheet);
            }

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

                DefaultSheet = selected[0].Sheet
            };

            Settings = settings;
            SelectedJobs = selected;

            DialogResult = DialogResult.OK;
            Close();
        }
    }

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
