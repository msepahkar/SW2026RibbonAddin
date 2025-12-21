using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nnesting";
        public string Tooltip => "Nest combined thickness DWGs into laser sheets (0/90/180/270 rotations).";
        public string Hint => "Laser cut nesting";

        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolder();
            if (string.IsNullOrEmpty(mainFolder))
                return;

            LaserCutRunSettings settings;
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
            }

            try
            {
                DwgLaserNester.NestFolder(mainFolder, settings, showUi: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Laser nesting failed:\r\n\r\n" + ex.Message,
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            return AddinContext.Enable;
        }

        private static string SelectMainFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder that contains thickness_*.dwg (combined outputs from Combine DWG)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }
    }

    internal readonly struct SheetPreset
    {
        public string Name { get; }
        public double WidthMm { get; }
        public double HeightMm { get; }

        public SheetPreset(string name, double widthMm, double heightMm)
        {
            Name = name ?? "";
            WidthMm = widthMm;
            HeightMm = heightMm;
        }

        public override string ToString()
        {
            return $"{Name} ({WidthMm:0.###} x {HeightMm:0.###} mm)";
        }
    }

    internal sealed class LaserCutRunSettings
    {
        public SheetPreset DefaultSheet { get; set; }

        // Exact material grouping (no normalization)
        public bool SeparateByMaterialExact { get; set; } = true;

        // If SeparateByMaterialExact is true, create one output DWG per material string
        public bool OutputOneDwgPerMaterial { get; set; } = true;

        // Option 2 behavior: keep only this material's source preview (plates + labels) in each output
        public bool KeepOnlyCurrentMaterialInSourcePreview { get; set; } = true;

        // Kept ONLY so older code (like BatchCombineNestButton) doesn't break if it prints these.
        // With exact materials, these presets are not used.
        public bool UsePerMaterialSheetPresets { get; set; } = false;
        public SheetPreset SteelSheet { get; set; }
        public SheetPreset AluminumSheet { get; set; }
        public SheetPreset StainlessSheet { get; set; }
        public SheetPreset OtherSheet { get; set; }
    }

    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly ComboBox _preset;
        private readonly NumericUpDown _w;
        private readonly NumericUpDown _h;

        private readonly CheckBox _sepMat;
        private readonly CheckBox _oneDwgPerMat;
        private readonly CheckBox _filterPreview;

        private readonly Button _ok;
        private readonly Button _cancel;

        private readonly List<SheetPreset> _presets = new List<SheetPreset>
        {
            new SheetPreset("1500 x 3000 mm", 3000, 1500),
            new SheetPreset("1250 x 2500 mm", 2500, 1250),
            new SheetPreset("1000 x 2000 mm", 2000, 1000),
            new SheetPreset("Custom", 3000, 1500),
        };

        public LaserCutRunSettings Settings { get; private set; }

        public LaserCutOptionsForm()
        {
            Text = "Batch nest options (applies to ALL thickness files)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Width = 560;
            Height = 270;

            var lblPreset = new Label { Left = 12, Top = 16, Width = 170, Text = "Global sheet preset:" };
            _preset = new ComboBox
            {
                Left = 190,
                Top = 12,
                Width = 330,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            foreach (var p in _presets)
                _preset.Items.Add(p.ToString());
            _preset.SelectedIndex = 0;
            _preset.SelectedIndexChanged += (_, __) => ApplyPresetToNumeric();

            var lblW = new Label { Left = 12, Top = 52, Width = 170, Text = "Width (mm):" };
            _w = new NumericUpDown
            {
                Left = 190,
                Top = 48,
                Width = 120,
                DecimalPlaces = 1,
                Minimum = 100,
                Maximum = 200000,
                Value = 3000
            };

            var lblH = new Label { Left = 320, Top = 52, Width = 80, Text = "Height:" };
            _h = new NumericUpDown
            {
                Left = 400,
                Top = 48,
                Width = 120,
                DecimalPlaces = 1,
                Minimum = 100,
                Maximum = 200000,
                Value = 1500
            };

            _sepMat = new CheckBox
            {
                Left = 12,
                Top = 86,
                Width = 520,
                Text = "Separate nests by EXACT SolidWorks material name (no normalization / no translation)",
                Checked = true
            };
            _sepMat.CheckedChanged += (_, __) =>
            {
                _oneDwgPerMat.Enabled = _sepMat.Checked;
                _filterPreview.Enabled = _sepMat.Checked;
                if (!_sepMat.Checked)
                {
                    _oneDwgPerMat.Checked = false;
                    _filterPreview.Checked = false;
                }
                else
                {
                    _oneDwgPerMat.Checked = true;
                    _filterPreview.Checked = true;
                }
            };

            _oneDwgPerMat = new CheckBox
            {
                Left = 32,
                Top = 112,
                Width = 520,
                Text = "Output one nested DWG per material name",
                Checked = true
            };

            _filterPreview = new CheckBox
            {
                Left = 32,
                Top = 136,
                Width = 520,
                Text = "Keep only that material's source preview (plates + labels) in each output (Option 2)",
                Checked = true
            };

            var note = new Label
            {
                Left = 12,
                Top = 162,
                Width = 520,
                Text = "Note: rotations are always 0/90/180/270. Gap+margin are auto (>= thickness)."
            };

            _ok = new Button { Text = "OK", Left = 360, Width = 75, Top = 190, DialogResult = DialogResult.OK };
            _cancel = new Button { Text = "Cancel", Left = 445, Width = 75, Top = 190, DialogResult = DialogResult.Cancel };

            AcceptButton = _ok;
            CancelButton = _cancel;

            Controls.Add(lblPreset);
            Controls.Add(_preset);
            Controls.Add(lblW);
            Controls.Add(_w);
            Controls.Add(lblH);
            Controls.Add(_h);

            Controls.Add(_sepMat);
            Controls.Add(_oneDwgPerMat);
            Controls.Add(_filterPreview);
            Controls.Add(note);

            Controls.Add(_ok);
            Controls.Add(_cancel);

            ApplyPresetToNumeric();

            // Build settings on OK
            _ok.Click += (_, __) =>
            {
                var chosen = _presets[Math.Max(0, _preset.SelectedIndex)];
                var final = new SheetPreset(
                    chosen.Name == "Custom" ? "Custom" : chosen.Name,
                    (double)_w.Value,
                    (double)_h.Value);

                Settings = new LaserCutRunSettings
                {
                    DefaultSheet = final,
                    SeparateByMaterialExact = _sepMat.Checked,
                    OutputOneDwgPerMaterial = _sepMat.Checked && _oneDwgPerMat.Checked,
                    KeepOnlyCurrentMaterialInSourcePreview = _sepMat.Checked && _filterPreview.Checked,

                    // compatibility fields:
                    UsePerMaterialSheetPresets = false
                };
            };
        }

        private void ApplyPresetToNumeric()
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
        private int _total;
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
            Width = 520;
            Height = 130;

            _label = new Label { Left = 12, Top = 12, Width = 480, Text = "Starting..." };
            _bar = new ProgressBar { Left = 12, Top = 40, Width = 480, Height = 20, Minimum = 0, Maximum = _total, Value = 0 };

            Controls.Add(_label);
            Controls.Add(_bar);
        }

        public void Step(string message)
        {
            _done++;
            if (_done > _total) _done = _total;

            _label.Text = message ?? "";
            _bar.Value = _done;

            // keep UI responsive
            System.Windows.Forms.Application.DoEvents();
        }
    }

    internal static class DwgLaserNester
    {
        internal sealed class NestRunResult
        {
            public string ThicknessFile;
            public string MaterialExact;
            public string OutputDwg;
            public int SheetsUsed;
            public int TotalParts;
        }

        private sealed class PartDefinition
        {
            public BlockRecord Block;
            public string BlockName;

            public string MaterialExact;

            public double MinX;
            public double MinY;
            public double MaxX;
            public double MaxY;

            public double Width;
            public double Height;

            public int Quantity;
        }

        private sealed class SheetState
        {
            public int Index;
            public double OriginX;
            public double OriginY;
            public List<FreeRect> FreeRects = new List<FreeRect>();
        }

        private sealed class FreeRect
        {
            public double X;
            public double Y;
            public double Width;
            public double Height;
        }

        public static void NestFolder(string mainFolder, LaserCutRunSettings settings, bool showUi = true)
        {
            if (settings.DefaultSheet.WidthMm <= 0 || settings.DefaultSheet.HeightMm <= 0)
                throw new ArgumentException("Sheet width/height must be positive.");

            if (string.IsNullOrWhiteSpace(mainFolder) || !Directory.Exists(mainFolder))
                throw new DirectoryNotFoundException("Folder not found: " + mainFolder);

            // Only take the combined thickness outputs, not already nested results
            var thicknessFiles = Directory.GetFiles(mainFolder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
                .Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f) ?? "";
                    return !n.Contains("_nested", StringComparison.OrdinalIgnoreCase)
                           && !n.Contains("_nest_", StringComparison.OrdinalIgnoreCase);
                })
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (thicknessFiles.Count == 0)
            {
                if (showUi)
                {
                    MessageBox.Show(
                        "No thickness_*.dwg files found in this folder.\r\n" +
                        "Run Combine DWG first (it creates thickness_*.dwg).",
                        "Laser nesting",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return;
            }

            var batchSummary = new StringBuilder();
            batchSummary.AppendLine("Batch nesting summary");
            batchSummary.AppendLine("Folder: " + mainFolder);
            batchSummary.AppendLine("Sheet: " + settings.DefaultSheet);
            batchSummary.AppendLine("SeparateByMaterialExact: " + settings.SeparateByMaterialExact);
            batchSummary.AppendLine("OutputOneDwgPerMaterial: " + settings.OutputOneDwgPerMaterial);
            batchSummary.AppendLine("KeepOnlyCurrentMaterialInSourcePreview: " + settings.KeepOnlyCurrentMaterialInSourcePreview);
            batchSummary.AppendLine(new string('-', 60));

            foreach (var thicknessFile in thicknessFiles)
            {
                var results = NestThicknessFile(thicknessFile, settings);

                batchSummary.AppendLine(Path.GetFileName(thicknessFile));
                foreach (var r in results)
                {
                    batchSummary.AppendLine($"  Material: {r.MaterialExact}");
                    batchSummary.AppendLine($"  SheetsUsed: {r.SheetsUsed}, Parts: {r.TotalParts}");
                    batchSummary.AppendLine($"  Output: {Path.GetFileName(r.OutputDwg)}");
                }
                batchSummary.AppendLine(new string('-', 60));
            }

            string summaryPath = Path.Combine(mainFolder, "batch_nest_summary.txt");
            File.WriteAllText(summaryPath, batchSummary.ToString(), Encoding.UTF8);

            if (showUi)
            {
                MessageBox.Show(
                    "Batch nesting finished.\r\n\r\nSummary:\r\n" + summaryPath,
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        // Back-compat overloads (in case some other button calls these)
        public static void Nest(string sourceDwgPath, double sheetWidth, double sheetHeight)
        {
            var settings = new LaserCutRunSettings
            {
                DefaultSheet = new SheetPreset("Custom", sheetWidth, sheetHeight),
                SeparateByMaterialExact = false,
                OutputOneDwgPerMaterial = false,
                KeepOnlyCurrentMaterialInSourcePreview = false
            };

            NestThicknessFile(sourceDwgPath, settings);
        }

        public static List<NestRunResult> Nest(string thicknessFile, LaserCutRunSettings settings)
        {
            return NestThicknessFile(thicknessFile, settings);
        }

        public static List<NestRunResult> NestThicknessFile(string sourceDwgPath, LaserCutRunSettings settings)
        {
            if (!File.Exists(sourceDwgPath))
                throw new FileNotFoundException("DWG file not found.", sourceDwgPath);

            // First pass: read once to find materials + quantities
            CadDocument doc0;
            using (var reader = new DwgReader(sourceDwgPath))
                doc0 = reader.Read();

            var defs0 = LoadPartDefinitions(doc0).ToList();
            if (defs0.Count == 0)
                throw new InvalidOperationException("No plate blocks (P_*_Q#) found in: " + sourceDwgPath);

            var groups = BuildGroups(defs0, settings);

            var results = new List<NestRunResult>();

            foreach (var grp in groups)
            {
                string materialKey = grp.Key;
                string materialLabel = grp.Value; // exact (pretty) label

                // Re-read fresh doc for each output (so outputs don't overlap each other)
                CadDocument doc;
                using (var reader = new DwgReader(sourceDwgPath))
                    doc = reader.Read();

                var defs = LoadPartDefinitions(doc)
                    .Where(d => MaterialNameCodec.Normalize(d.MaterialExact)
                        .Equals(materialKey, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                int totalInstances = defs.Sum(d => d.Quantity);
                if (totalInstances <= 0)
                    continue;

                // Determine auto gap+margin (>= thickness)
                double thicknessMm = TryGetPlateThicknessFromFileName(sourceDwgPath) ?? 0.0;

                double partGap = 3.0;
                if (thicknessMm > partGap)
                    partGap = thicknessMm;

                double sheetMargin = 10.0;
                if (thicknessMm > sheetMargin)
                    sheetMargin = thicknessMm;

                double placementMargin = sheetMargin + 2.0 * partGap;

                double sheetWidth = settings.DefaultSheet.WidthMm;
                double sheetHeight = settings.DefaultSheet.HeightMm;

                // Validate fit
                double usableWidth = sheetWidth - 2 * placementMargin;
                double usableHeight = sheetHeight - 2 * placementMargin;

                foreach (var d in defs)
                {
                    if (d.Width > usableWidth || d.Height > usableHeight)
                    {
                        throw new InvalidOperationException(
                            $"Part '{d.BlockName}' ({d.Width:0.##} x {d.Height:0.##} mm) " +
                            $"does not fit inside sheet {sheetWidth:0.##} x {sheetHeight:0.##} mm " +
                            $"with margin {placementMargin:0.##} mm.");
                    }
                }

                // Option 2: keep only that material’s source preview
                if (settings.SeparateByMaterialExact &&
                    settings.OutputOneDwgPerMaterial &&
                    settings.KeepOnlyCurrentMaterialInSourcePreview)
                {
                    FilterSourcePreviewToTheseBlocks(doc, defs.Select(d => d.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase));
                }

                // Expand instances
                var instances = new List<PartDefinition>(totalInstances);
                foreach (var d in defs)
                {
                    for (int i = 0; i < d.Quantity; i++)
                        instances.Add(d);
                }

                // Sort biggest first
                instances.Sort((a, b) =>
                {
                    double areaA = a.Width * a.Height;
                    double areaB = b.Width * b.Height;
                    return areaB.CompareTo(areaA);
                });

                var modelSpace = doc.BlockRecords["*Model_Space"];

                // Compute extents of remaining source preview
                GetModelSpaceExtents(doc, out double origMinX, out double origMinY, out double origMaxX, out double origMaxY);

                // Place sheets ABOVE source preview
                double baseSheetOriginY = origMaxY + 200.0;
                double baseSheetOriginX = origMinX;

                using (var progress = new LaserCutProgressForm(totalInstances))
                {
                    progress.Show();
                    System.Windows.Forms.Application.DoEvents();

                    int sheetCount = NestFreeRectangles(
                        instances,
                        modelSpace,
                        sheetWidth,
                        sheetHeight,
                        placementMargin,
                        sheetGap: 50.0,
                        partGap,
                        baseSheetOriginX,
                        baseSheetOriginY,
                        progress,
                        totalInstances,
                        materialLabel);

                    // Write output
                    string dir = Path.GetDirectoryName(sourceDwgPath) ?? "";
                    string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath) ?? "thickness";

                    string outPath;

                    if (settings.SeparateByMaterialExact && settings.OutputOneDwgPerMaterial)
                    {
                        string safeMat = MaterialNameCodec.MakeSafeFileToken(materialLabel);
                        outPath = Path.Combine(dir, $"{nameNoExt}_nested_{safeMat}.dwg");
                    }
                    else
                    {
                        outPath = Path.Combine(dir, $"{nameNoExt}_nested.dwg");
                    }

                    using (var writer = new DwgWriter(outPath, doc))
                        writer.Write();

                    // Write a small log next to output
                    string logPath = Path.Combine(dir, $"{nameNoExt}_nest_log.txt");
                    AppendNestLog(logPath, sourceDwgPath, materialLabel, sheetWidth, sheetHeight, thicknessMm, partGap, sheetCount, totalInstances, outPath);

                    results.Add(new NestRunResult
                    {
                        ThicknessFile = sourceDwgPath,
                        MaterialExact = materialLabel,
                        OutputDwg = outPath,
                        SheetsUsed = sheetCount,
                        TotalParts = totalInstances
                    });

                    progress.Close();
                }
            }

            return results;
        }

        private static void AppendNestLog(
            string logPath,
            string thicknessFile,
            string material,
            double sheetW,
            double sheetH,
            double thicknessMm,
            double gapMm,
            int sheets,
            int parts,
            string outDwg)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Nest run:");
                sb.AppendLine("  Thickness file: " + Path.GetFileName(thicknessFile));
                sb.AppendLine("  Material: " + material);
                sb.AppendLine($"  Sheet: {sheetW:0.###} x {sheetH:0.###} mm");
                sb.AppendLine($"  Thickness(mm): {thicknessMm:0.###}");
                sb.AppendLine($"  Gap(mm): {gapMm:0.###}  (auto >= thickness)");
                sb.AppendLine($"  Sheets used: {sheets}");
                sb.AppendLine($"  Total parts: {parts}");
                sb.AppendLine("  Output: " + Path.GetFileName(outDwg));
                sb.AppendLine(new string('-', 60));

                File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // ignore logging failures
            }
        }

        private static Dictionary<string, string> BuildGroups(List<PartDefinition> defs, LaserCutRunSettings settings)
        {
            // key = normalized string used for grouping (case-insensitive compare)
            // value = the exact (pretty) label we will show / use in outputs
            var groups = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!settings.SeparateByMaterialExact || !settings.OutputOneDwgPerMaterial)
            {
                groups["ALL"] = "ALL";
                return groups;
            }

            foreach (var d in defs)
            {
                string mat = MaterialNameCodec.Normalize(d.MaterialExact);
                if (!groups.ContainsKey(mat))
                    groups[mat] = mat;
            }

            if (groups.Count == 0)
                groups["UNKNOWN"] = "UNKNOWN";

            return groups;
        }

        private static double? TryGetPlateThicknessFromFileName(string sourceDwgPath)
        {
            if (string.IsNullOrWhiteSpace(sourceDwgPath))
                return null;

            string fileName = Path.GetFileNameWithoutExtension(sourceDwgPath);
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            const string prefix = "thickness_";
            int idx = fileName.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return null;

            string token = fileName.Substring(idx + prefix.Length);
            if (string.IsNullOrWhiteSpace(token))
                return null;

            token = token.Replace('_', '.');

            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                value > 0.0 && value < 1000.0)
            {
                return value;
            }

            return null;
        }

        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc)
        {
            if (doc == null)
                yield break;

            foreach (var br in doc.BlockRecords)
            {
                var block = br;
                if (block == null)
                    continue;

                string name = block.Name ?? "";
                if (!name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                int qIndex = name.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
                if (qIndex < 0)
                    continue;

                int qty = 1;
                string qtyToken = name.Substring(qIndex + 2);
                if (!int.TryParse(qtyToken, NumberStyles.Integer, CultureInfo.InvariantCulture, out qty))
                    qty = 1;

                // material exact from token in block name
                string material = "UNKNOWN";
                MaterialNameCodec.TryExtractFromBlockName(name, out material);

                // bbox of entities inside the block (local coords)
                double minX = double.MaxValue, minY = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue;

                foreach (var ent in block.Entities)
                {
                    try
                    {
                        var bb = ent.GetBoundingBox();
                        XYZ bbMin = bb.Min;
                        XYZ bbMax = bb.Max;

                        if (bbMin.X < minX) minX = bbMin.X;
                        if (bbMin.Y < minY) minY = bbMin.Y;
                        if (bbMax.X > maxX) maxX = bbMax.X;
                        if (bbMax.Y > maxY) maxY = bbMax.Y;
                    }
                    catch
                    {
                        // ignore entities without bbox
                    }
                }

                if (minX == double.MaxValue || maxX == double.MinValue)
                    continue;

                double w = maxX - minX;
                double h = maxY - minY;

                if (w <= 0.0 || h <= 0.0)
                    continue;

                yield return new PartDefinition
                {
                    Block = block,
                    BlockName = name,
                    MaterialExact = material,
                    MinX = minX,
                    MinY = minY,
                    MaxX = maxX,
                    MaxY = maxY,
                    Width = w,
                    Height = h,
                    Quantity = Math.Max(1, qty)
                };
            }
        }

        private static void FilterSourcePreviewToTheseBlocks(CadDocument doc, HashSet<string> keepBlockNames)
        {
            if (doc == null || keepBlockNames == null || keepBlockNames.Count == 0)
                return;

            BlockRecord modelSpace;
            try
            {
                modelSpace = doc.BlockRecords["*Model_Space"];
            }
            catch
            {
                return;
            }

            // Remove inserts of plate blocks that are NOT in the keep set.
            // Keep MTexts only if their X falls near any kept insert bbox.
            var allInserts = modelSpace.Entities.OfType<Insert>().ToList();

            // Build world X-ranges for kept inserts
            var keepXRanges = new List<(double minX, double maxX)>();

            foreach (var ins in allInserts)
            {
                var blk = ins?.Block;
                if (blk == null) continue;

                string bn = blk.Name ?? "";
                if (!bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!keepBlockNames.Contains(bn))
                    continue;

                // Combined layout inserts are non-rotated; bbox in world = insert + local bbox
                // We approximate by using the block's bbox:
                var def = LoadPartDefinitions(doc).FirstOrDefault(d => d.BlockName.Equals(bn, StringComparison.OrdinalIgnoreCase));
                if (def == null) continue;

                double ix = ins.InsertPoint.X;
                double minX = ix + def.MinX;
                double maxX = ix + def.MaxX;
                if (minX > maxX) { var t = minX; minX = maxX; maxX = t; }

                keepXRanges.Add((minX, maxX));
            }

            // Pad so we keep labels that are slightly wider than geometry
            const double pad = 80.0;

            bool IsNearKeptX(double x)
            {
                foreach (var r in keepXRanges)
                {
                    if (x >= r.minX - pad && x <= r.maxX + pad)
                        return true;
                }
                return false;
            }

            var toRemove = new List<Entity>();

            foreach (var ent in modelSpace.Entities)
            {
                if (ent is Insert ins)
                {
                    var blk = ins.Block;
                    if (blk == null) continue;

                    string bn = blk.Name ?? "";
                    if (bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase) && !keepBlockNames.Contains(bn))
                        toRemove.Add(ent);
                }
            }

            foreach (var ent in modelSpace.Entities)
            {
                if (ent is MText mt)
                {
                    double x = mt.InsertPoint.X;
                    if (!IsNearKeptX(x))
                        toRemove.Add(ent);
                }
            }

            foreach (var ent in toRemove.Distinct())
            {
                try { modelSpace.Entities.Remove(ent); } catch { }
            }
        }

        private static void GetModelSpaceExtents(CadDocument doc, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            BlockRecord modelSpace;
            try
            {
                modelSpace = doc.BlockRecords["*Model_Space"];
            }
            catch
            {
                minX = minY = maxX = maxY = 0.0;
                return;
            }

            foreach (var ent in modelSpace.Entities)
            {
                try
                {
                    var bb = ent.GetBoundingBox();
                    XYZ a = bb.Min;
                    XYZ b = bb.Max;

                    if (a.X < minX) minX = a.X;
                    if (a.Y < minY) minY = a.Y;
                    if (b.X > maxX) maxX = b.X;
                    if (b.Y > maxY) maxY = b.Y;
                }
                catch
                {
                    // ignore
                }
            }

            if (minX == double.MaxValue || maxX == double.MinValue)
            {
                minX = minY = maxX = maxY = 0.0;
            }
        }

        private static int NestFreeRectangles(
            List<PartDefinition> instances,
            BlockRecord modelSpace,
            double sheetWidth,
            double sheetHeight,
            double placementMargin,
            double sheetGap,
            double partGap,
            double startOriginX,
            double baseOriginY,
            LaserCutProgressForm progress,
            int totalInstances,
            string materialLabel)
        {
            var sheets = new List<SheetState>();

            SheetState NewSheet()
            {
                var sheet = new SheetState
                {
                    Index = sheets.Count + 1,
                    OriginX = startOriginX + sheets.Count * (sheetWidth + sheetGap),
                    OriginY = baseOriginY
                };

                sheets.Add(sheet);
                DrawSheetOutline(sheet, sheetWidth, sheetHeight, modelSpace, materialLabel);

                sheet.FreeRects.Add(new FreeRect
                {
                    X = placementMargin,
                    Y = placementMargin,
                    Width = sheetWidth - 2 * placementMargin,
                    Height = sheetHeight - 2 * placementMargin
                });

                return sheet;
            }

            int placed = 0;
            var sheetState = NewSheet();

            foreach (var inst in instances)
            {
                while (true)
                {
                    if (TryPlaceOnSheet(sheetState, inst, partGap, modelSpace, ref placed, totalInstances, progress))
                        break;

                    sheetState = NewSheet();
                }
            }

            return sheets.Count;
        }

        private static bool TryPlaceOnSheet(
            SheetState sheet,
            PartDefinition part,
            double partGap,
            BlockRecord modelSpace,
            ref int placed,
            int totalInstances,
            LaserCutProgressForm progress)
        {
            if (sheet.FreeRects.Count == 0)
                return false;

            int bestIndex = -1;
            bool bestRot90 = false;
            double bestShortSideFit = double.MaxValue;
            double bestLongSideFit = double.MaxValue;

            for (int i = 0; i < sheet.FreeRects.Count; i++)
            {
                var fr = sheet.FreeRects[i];

                Evaluate(fr, part.Width, part.Height, false, i);
                Evaluate(fr, part.Height, part.Width, true, i);
            }

            if (bestIndex < 0)
                return false;

            var chosen = sheet.FreeRects[bestIndex];
            bool rot90 = bestRot90;

            double w = rot90 ? part.Height : part.Width;
            double h = rot90 ? part.Width : part.Height;

            double placeW = w + partGap;
            double placeH = h + partGap;

            double desiredMinLocalX = chosen.X + partGap * 0.5;
            double desiredMinLocalY = chosen.Y + partGap * 0.5;

            double insertXWorld;
            double insertYWorld;
            double rotationRad;

            if (!rot90)
            {
                double worldMinX = sheet.OriginX + desiredMinLocalX;
                double worldMinY = sheet.OriginY + desiredMinLocalY;

                insertXWorld = worldMinX - part.MinX;
                insertYWorld = worldMinY - part.MinY;

                // 0 or 180 are equivalent for packing; we keep 0.
                rotationRad = 0.0;
            }
            else
            {
                // 90° rotation around insert point:
                // rotate (x,y)->(-y,x).
                // bounding min after rotation about origin: (-MaxY, MinX)
                double worldMinX = sheet.OriginX + desiredMinLocalX;
                double worldMinY = sheet.OriginY + desiredMinLocalY;

                insertXWorld = worldMinX + part.MaxY;
                insertYWorld = worldMinY - part.MinX;

                // 90 or 270 are equivalent for packing; we keep 90.
                rotationRad = Math.PI / 2.0;
            }

            var insert = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insertXWorld, insertYWorld, 0.0),
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0,
                Rotation = rotationRad
            };
            modelSpace.Entities.Add(insert);

            SplitFreeRect(sheet, bestIndex, chosen, placeW, placeH);

            placed++;
            progress.Step($"Placed {placed} / {totalInstances} ...");

            return true;

            void Evaluate(FreeRect fr, double wc, double hc, bool candidateRot90, int rectIndex)
            {
                double pw = wc + partGap;
                double ph = hc + partGap;

                if (pw > fr.Width || ph > fr.Height)
                    return;

                double leftoverHoriz = fr.Width - pw;
                double leftoverVert = fr.Height - ph;

                double shortSideFit = Math.Min(leftoverHoriz, leftoverVert);
                double longSideFit = Math.Max(leftoverHoriz, leftoverVert);

                if (shortSideFit < bestShortSideFit ||
                    (Math.Abs(shortSideFit - bestShortSideFit) < 1e-9 && longSideFit < bestLongSideFit))
                {
                    bestShortSideFit = shortSideFit;
                    bestLongSideFit = longSideFit;
                    bestIndex = rectIndex;
                    bestRot90 = candidateRot90;
                }
            }
        }

        private static void SplitFreeRect(SheetState sheet, int rectIndex, FreeRect usedRect, double usedWidth, double usedHeight)
        {
            sheet.FreeRects.RemoveAt(rectIndex);

            const double minSize = 1.0;

            double rightWidth = usedRect.Width - usedWidth;
            if (rightWidth > minSize)
            {
                sheet.FreeRects.Add(new FreeRect
                {
                    X = usedRect.X + usedWidth,
                    Y = usedRect.Y,
                    Width = rightWidth,
                    Height = usedRect.Height
                });
            }

            double topHeight = usedRect.Height - usedHeight;
            if (topHeight > minSize)
            {
                sheet.FreeRects.Add(new FreeRect
                {
                    X = usedRect.X,
                    Y = usedRect.Y + usedHeight,
                    Width = usedWidth,
                    Height = topHeight
                });
            }
        }

        private static void DrawSheetOutline(SheetState sheet, double sheetWidth, double sheetHeight, BlockRecord modelSpace, string materialLabel)
        {
            // Border
            modelSpace.Entities.Add(new Line
            {
                StartPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0),
                EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0)
            });
            modelSpace.Entities.Add(new Line
            {
                StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0),
                EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0)
            });
            modelSpace.Entities.Add(new Line
            {
                StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0),
                EndPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0)
            });
            modelSpace.Entities.Add(new Line
            {
                StartPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0),
                EndPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0)
            });

            // Title (material)
            if (!string.IsNullOrWhiteSpace(materialLabel) && !materialLabel.Equals("ALL", StringComparison.OrdinalIgnoreCase))
            {
                modelSpace.Entities.Add(new MText
                {
                    Value = $"Material: {materialLabel}",
                    InsertPoint = new XYZ(sheet.OriginX + 10.0, sheet.OriginY + sheetHeight + 15.0, 0.0),
                    Height = 20.0
                });
            }
        }
    }
}
