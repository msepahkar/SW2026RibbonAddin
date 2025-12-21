// Commands\dwg\LaserCutButton.cs
// DROP-IN: Replace the entire file with this.
// Requires NuGet: Clipper2Lib

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;

using CSMath;           // XYZ
using Clipper2Lib;      // Clipper64, Paths64, Path64, Point64, ClipperOffset, etc.

using WinPoint = System.Drawing.Point;
using WinSize = System.Drawing.Size;
using WinContentAlignment = System.Drawing.ContentAlignment;
using ClipperClipType = Clipper2Lib.ClipType;


namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nnesting";
        public string Tooltip => "Nest combined thickness DWGs into sheets. Supports Fast (rectangles) or Contour (Level 1).";
        public string Hint => "Laser cut nesting";

        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string folder = SelectFolder();
            if (string.IsNullOrWhiteSpace(folder))
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
                DwgLaserNester.NestFolder(folder, settings, showUi: true);
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

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        private static string SelectFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the folder that contains thickness_*.dwg (outputs of Combine DWG)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }
    }

    internal enum NestingMode
    {
        FastRectangles = 0,
        ContourLevel1 = 1,
    }

    internal readonly struct SheetPreset
    {
        public string Name { get; }
        public double WidthMm { get; }
        public double HeightMm { get; }

        public SheetPreset(string name, double wMm, double hMm)
        {
            Name = name ?? "";
            WidthMm = wMm;
            HeightMm = hMm;
        }

        public override string ToString() => $"{Name} ({WidthMm:0.###} x {HeightMm:0.###} mm)";
    }

    internal sealed class LaserCutRunSettings
    {
        public SheetPreset DefaultSheet { get; set; }

        // EXACT material grouping (no normalization)
        public bool SeparateByMaterialExact { get; set; } = true;

        // If SeparateByMaterialExact is true, create one output DWG per material string
        public bool OutputOneDwgPerMaterial { get; set; } = true;

        // Option 2 behavior: keep only this material's preview in each output
        public bool KeepOnlyCurrentMaterialInSourcePreview { get; set; } = true;

        // Nesting algorithm selection
        public NestingMode Mode { get; set; } = NestingMode.ContourLevel1;

        // Contour extraction tuning (mm)
        // Smaller chord => better contour, slower.
        public double ContourChordMm { get; set; } = 0.8;

        // Endpoint snapping tolerance for loop building (mm)
        public double ContourSnapMm { get; set; } = 0.05;

        // Candidate limit per (sheet,rotation) placement attempt (safety)
        public int MaxCandidatesPerTry { get; set; } = 6000;
    }

    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly ComboBox _preset;
        private readonly NumericUpDown _w;
        private readonly NumericUpDown _h;

        private readonly CheckBox _sepMat;
        private readonly CheckBox _onePerMat;
        private readonly CheckBox _filterPreview;

        private readonly RadioButton _rbFast;
        private readonly RadioButton _rbContour;

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

            ClientSize = new WinSize(640, 360);

            Controls.Add(new Label { Left = 12, Top = 16, Width = 170, Text = "Sheet preset:" });
            _preset = new ComboBox
            {
                Left = 180,
                Top = 12,
                Width = 440,
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
                Width = 610,
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
                Width = 610,
                Text = "Output one nested DWG per material",
                Checked = true
            };
            Controls.Add(_onePerMat);

            _filterPreview = new CheckBox
            {
                Left = 32,
                Top = 136,
                Width = 610,
                Text = "Keep only that material's source preview (plates + labels) in each output",
                Checked = true
            };
            Controls.Add(_filterPreview);

            var grp = new GroupBox
            {
                Left = 12,
                Top = 170,
                Width = 610,
                Height = 120,
                Text = "Nesting algorithm"
            };
            Controls.Add(grp);

            _rbFast = new RadioButton
            {
                Left = 16,
                Top = 24,
                Width = 560,
                Text = "Fast (Rectangles) — very fast, more wasted sheet",
                Checked = false
            };
            grp.Controls.Add(_rbFast);

            _rbContour = new RadioButton
            {
                Left = 16,
                Top = 48,
                Width = 560,
                Text = "Contour (Level 1) — uses outer contour + gap offset, better packing (slower)",
                Checked = true
            };
            grp.Controls.Add(_rbContour);

            grp.Controls.Add(new Label { Left = 36, Top = 78, Width = 160, Text = "Arc chord (mm):" });
            _chord = new NumericUpDown
            {
                Left = 200,
                Top = 74,
                Width = 90,
                DecimalPlaces = 2,
                Minimum = 0.10M,
                Maximum = 5.00M,
                Value = 0.80M
            };
            grp.Controls.Add(_chord);

            grp.Controls.Add(new Label { Left = 320, Top = 78, Width = 160, Text = "Snap tol (mm):" });
            _snap = new NumericUpDown
            {
                Left = 460,
                Top = 74,
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
                Top = 300,
                Width = 610,
                Height = 24,
                Text = "Note: rotations are always 0/90/180/270. Gap + margin are auto (>= thickness)."
            };
            Controls.Add(note);

            _ok = new Button { Text = "OK", Left = 450, Top = 328, Width = 80, Height = 26 };
            _cancel = new Button { Text = "Cancel", Left = 540, Top = 328, Width = 80, Height = 26, DialogResult = DialogResult.Cancel };
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

                Settings = new LaserCutRunSettings
                {
                    DefaultSheet = sheet,

                    SeparateByMaterialExact = _sepMat.Checked,
                    OutputOneDwgPerMaterial = _sepMat.Checked && _onePerMat.Checked,
                    KeepOnlyCurrentMaterialInSourcePreview = _sepMat.Checked && _filterPreview.Checked,

                    Mode = _rbContour.Checked ? NestingMode.ContourLevel1 : NestingMode.FastRectangles,
                    ContourChordMm = (double)_chord.Value,
                    ContourSnapMm = (double)_snap.Value,
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

            ClientSize = new WinSize(560, 95);

            _label = new Label
            {
                Left = 12,
                Top = 10,
                Width = 536,
                Height = 22,
                TextAlign = WinContentAlignment.MiddleLeft,
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
    }

    internal static class DwgLaserNester
    {
        // Geometry scale for Clipper (mm -> integer)
        private const long SCALE = 1000; // 0.001 mm units

        private static readonly int[] RotationsDeg = { 0, 90, 180, 270 };

        internal sealed class NestRunResult
        {
            public string ThicknessFile;
            public string MaterialExact;
            public string OutputDwg;
            public int SheetsUsed;
            public int TotalParts;
            public NestingMode Mode;
        }

        private sealed class PartDefinition
        {
            public BlockRecord Block;
            public string BlockName;
            public int Quantity;
            public string MaterialExact;

            // bbox in mm (used by Fast mode and also for fallback)
            public double MinX, MinY, MaxX, MaxY;
            public double Width, Height;

            // Contour polygon (outer boundary) in scaled Clipper units, in block-local coords (unrotated)
            public Path64 OuterContour0;
            public long OuterArea2Abs; // abs(2*area) in scaled^2
        }

        private sealed class FreeRect
        {
            public double X, Y, W, H;
        }

        private sealed class SheetRectState
        {
            public int Index;
            public double OriginXmm;
            public double OriginYmm;
            public List<FreeRect> Free = new List<FreeRect>();
            public int PlacedCount;
            public double UsedArea;
        }

        private sealed class SheetContourState
        {
            public int Index;
            public double OriginXmm;
            public double OriginYmm;

            public List<PlacedContour> Placed = new List<PlacedContour>();
            public int PlacedCount;
            public long UsedArea2Abs; // scaled^2
        }

        private sealed class PlacedContour
        {
            public Path64 OffsetPoly; // translated into sheet-local usable coords (scaled)
            public LongRect BBox;      // bbox of OffsetPoly (scaled)
        }

        private struct LongRect
        {
            public long MinX, MinY, MaxX, MaxY;
        }

        public static void NestFolder(string mainFolder, LaserCutRunSettings settings, bool showUi = true)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            if (string.IsNullOrWhiteSpace(mainFolder) || !Directory.Exists(mainFolder))
                throw new DirectoryNotFoundException("Folder not found: " + mainFolder);

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
                        "No thickness_*.dwg files found in this folder.\r\nRun Combine DWG first.",
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
            batchSummary.AppendLine("Mode: " + settings.Mode);
            batchSummary.AppendLine(new string('-', 70));

            foreach (var thicknessFile in thicknessFiles)
            {
                var results = NestThicknessFile(thicknessFile, settings);

                batchSummary.AppendLine(Path.GetFileName(thicknessFile));
                foreach (var r in results)
                {
                    batchSummary.AppendLine($"  Material: {r.MaterialExact}");
                    batchSummary.AppendLine($"  Mode: {r.Mode}");
                    batchSummary.AppendLine($"  SheetsUsed: {r.SheetsUsed}, Parts: {r.TotalParts}");
                    batchSummary.AppendLine($"  Output: {Path.GetFileName(r.OutputDwg)}");
                }
                batchSummary.AppendLine(new string('-', 70));
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

        public static List<NestRunResult> NestThicknessFile(string sourceDwgPath, LaserCutRunSettings settings)
        {
            if (!File.Exists(sourceDwgPath))
                throw new FileNotFoundException("DWG file not found.", sourceDwgPath);

            CadDocument firstDoc;
            using (var reader = new DwgReader(sourceDwgPath))
                firstDoc = reader.Read();

            var defsFirst = LoadPartDefinitions(firstDoc, settings).ToList();
            if (defsFirst.Count == 0)
                throw new InvalidOperationException("No plate blocks (P_*_Q#) found in: " + sourceDwgPath);

            // group by exact material
            var groups = BuildGroups(defsFirst, settings);

            var results = new List<NestRunResult>();

            foreach (var grp in groups)
            {
                string groupKey = grp.Key;       // normalized key
                string groupLabel = grp.Value;   // exact label

                // Re-read fresh doc for each output
                CadDocument doc;
                using (var reader = new DwgReader(sourceDwgPath))
                    doc = reader.Read();

                var defs = LoadPartDefinitions(doc, settings)
                    .Where(d => NormalizeKey(d.MaterialExact).Equals(groupKey, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                int totalInstances = defs.Sum(d => d.Quantity);
                if (totalInstances <= 0)
                    continue;

                // auto gap + margin
                double thicknessMm = TryGetPlateThicknessFromFileName(sourceDwgPath) ?? 0.0;

                double gapMm = 3.0;
                if (thicknessMm > gapMm) gapMm = thicknessMm;

                double marginMm = 10.0;
                if (thicknessMm > marginMm) marginMm = thicknessMm;

                var modelSpace = doc.BlockRecords["*Model_Space"];

                // Option 2: keep only material preview
                if (settings.SeparateByMaterialExact &&
                    settings.OutputOneDwgPerMaterial &&
                    settings.KeepOnlyCurrentMaterialInSourcePreview &&
                    !string.Equals(groupLabel, "ALL", StringComparison.OrdinalIgnoreCase))
                {
                    FilterSourcePreviewToTheseBlocks(doc, defs.Select(d => d.BlockName).ToHashSet(StringComparer.OrdinalIgnoreCase));
                }

                // Find extents of remaining preview to place sheets above it
                GetModelSpaceExtents(doc, out double srcMinX, out double srcMinY, out double srcMaxX, out double srcMaxY);

                double baseSheetOriginX = srcMinX;
                double baseSheetOriginY = srcMaxY + 200.0;

                // Run nesting with progress
                using (var progress = new LaserCutProgressForm(totalInstances))
                {
                    progress.Show();
                    Application.DoEvents();

                    int sheetsUsed;

                    if (settings.Mode == NestingMode.FastRectangles)
                    {
                        sheetsUsed = NestFastRectangles(
                            defs,
                            modelSpace,
                            settings.DefaultSheet.WidthMm,
                            settings.DefaultSheet.HeightMm,
                            marginMm,
                            gapMm,
                            baseSheetOriginX,
                            baseSheetOriginY,
                            groupLabel,
                            progress,
                            totalInstances);
                    }
                    else
                    {
                        sheetsUsed = NestContourLevel1(
                            defs,
                            modelSpace,
                            settings.DefaultSheet.WidthMm,
                            settings.DefaultSheet.HeightMm,
                            marginMm,
                            gapMm,
                            baseSheetOriginX,
                            baseSheetOriginY,
                            groupLabel,
                            progress,
                            totalInstances,
                            chordMm: Math.Max(0.10, settings.ContourChordMm),
                            snapMm: Math.Max(0.01, settings.ContourSnapMm),
                            maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry));
                    }

                    progress.Close();

                    // Save
                    string dir = Path.GetDirectoryName(sourceDwgPath) ?? "";
                    string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath) ?? "thickness";

                    string outPath;
                    if (settings.SeparateByMaterialExact && settings.OutputOneDwgPerMaterial)
                    {
                        string safeMat = MaterialNameCodec.MakeSafeFileToken(groupLabel);
                        outPath = Path.Combine(dir, $"{nameNoExt}_nested_{safeMat}.dwg");
                    }
                    else
                    {
                        outPath = Path.Combine(dir, $"{nameNoExt}_nested.dwg");
                    }

                    using (var writer = new DwgWriter(outPath, doc))
                        writer.Write();

                    // Log
                    string logPath = Path.Combine(dir, $"{nameNoExt}_nest_log.txt");
                    AppendNestLog(
                        logPath,
                        sourceDwgPath,
                        groupLabel,
                        settings.DefaultSheet.WidthMm,
                        settings.DefaultSheet.HeightMm,
                        thicknessMm,
                        gapMm,
                        marginMm,
                        sheetsUsed,
                        totalInstances,
                        outPath,
                        settings.Mode);

                    results.Add(new NestRunResult
                    {
                        ThicknessFile = sourceDwgPath,
                        MaterialExact = groupLabel,
                        OutputDwg = outPath,
                        SheetsUsed = sheetsUsed,
                        TotalParts = totalInstances,
                        Mode = settings.Mode
                    });
                }
            }

            return results;
        }

        // -----------------------------------------
        // FAST MODE (Rectangles) - existing approach
        // -----------------------------------------
        private static int NestFastRectangles(
            List<PartDefinition> defs,
            BlockRecord modelSpace,
            double sheetWmm,
            double sheetHmm,
            double marginMm,
            double gapMm,
            double baseOriginXmm,
            double baseOriginYmm,
            string materialLabel,
            LaserCutProgressForm progress,
            int totalInstances)
        {
            // We keep the old conservative rule for fast mode
            double placementMargin = marginMm + gapMm;

            double usableW = sheetWmm - 2 * placementMargin;
            double usableH = sheetHmm - 2 * placementMargin;
            if (usableW <= 0 || usableH <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            // Expand instances
            var instances = new List<PartDefinition>();
            foreach (var d in defs)
                for (int i = 0; i < d.Quantity; i++)
                    instances.Add(d);

            // Sort by bbox area
            instances.Sort((a, b) => (b.Width * b.Height).CompareTo(a.Width * a.Height));

            var sheets = new List<SheetRectState>();

            SheetRectState NewSheet()
            {
                var s = new SheetRectState
                {
                    Index = sheets.Count + 1,
                    OriginXmm = baseOriginXmm + (sheets.Count) * (sheetWmm + 60.0),
                    OriginYmm = baseOriginYmm
                };

                DrawSheetOutline(s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, modelSpace, materialLabel, s.Index, NestingMode.FastRectangles);

                s.Free.Add(new FreeRect
                {
                    X = placementMargin,
                    Y = placementMargin,
                    W = sheetWmm - 2 * placementMargin,
                    H = sheetHmm - 2 * placementMargin
                });

                sheets.Add(s);
                return s;
            }

            var cur = NewSheet();

            int placed = 0;

            foreach (var part in instances)
            {
                while (true)
                {
                    if (TryPlaceRect(cur, part, gapMm, modelSpace))
                    {
                        placed++;
                        progress.Step($"Placed {placed}/{totalInstances} (Fast rectangles)");
                        break;
                    }

                    cur = NewSheet();
                }
            }

            return sheets.Count;
        }

        private static bool TryPlaceRect(SheetRectState sheet, PartDefinition part, double gapMm, BlockRecord modelSpace)
        {
            // Only 0/90 rotations matter for rectangles (180=0, 270=90)
            // But we still respect DWG by inserting unrotated (fast mode)
            for (int frIndex = 0; frIndex < sheet.Free.Count; frIndex++)
            {
                var fr = sheet.Free[frIndex];

                // Try both orientations
                if (TryPlaceRectOrientation(sheet, part, frIndex, fr, part.Width, part.Height, 0.0, gapMm, modelSpace))
                    return true;

                if (TryPlaceRectOrientation(sheet, part, frIndex, fr, part.Height, part.Width, Math.PI / 2.0, gapMm, modelSpace))
                    return true;
            }

            return false;
        }

        private static bool TryPlaceRectOrientation(
            SheetRectState sheet,
            PartDefinition part,
            int frIndex,
            FreeRect fr,
            double w,
            double h,
            double rotRad,
            double gapMm,
            BlockRecord modelSpace)
        {
            double usedW = w + gapMm;
            double usedH = h + gapMm;

            if (usedW > fr.W || usedH > fr.H)
                return false;

            double localMinX = fr.X + gapMm * 0.5;
            double localMinY = fr.Y + gapMm * 0.5;

            double worldMinX = sheet.OriginXmm + localMinX;
            double worldMinY = sheet.OriginYmm + localMinY;

            // For fast mode we use bbox-based insert computation
            double insertX;
            double insertY;

            if (Math.Abs(rotRad) < 1e-9)
            {
                insertX = worldMinX - part.MinX;
                insertY = worldMinY - part.MinY;
            }
            else
            {
                // 90° about origin: (x,y)->(-y,x) => bbox min is (-MaxY, MinX)
                insertX = worldMinX + part.MaxY;
                insertY = worldMinY - part.MinX;
            }

            var ins = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insertX, insertY, 0.0),
                Rotation = rotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            };
            modelSpace.Entities.Add(ins);

            // Split free rect (simple guillotine split)
            sheet.Free.RemoveAt(frIndex);

            double rightW = fr.W - usedW;
            if (rightW > 1.0)
                sheet.Free.Add(new FreeRect { X = fr.X + usedW, Y = fr.Y, W = rightW, H = fr.H });

            double topH = fr.H - usedH;
            if (topH > 1.0)
                sheet.Free.Add(new FreeRect { X = fr.X, Y = fr.Y + usedH, W = usedW, H = topH });

            return true;
        }

        // -----------------------------------------
        // CONTOUR MODE (Level 1) - polygon packing
        // -----------------------------------------
        private static int NestContourLevel1(
            List<PartDefinition> defs,
            BlockRecord modelSpace,
            double sheetWmm,
            double sheetHmm,
            double marginMm,
            double gapMm,
            double baseOriginXmm,
            double baseOriginYmm,
            string materialLabel,
            LaserCutProgressForm progress,
            int totalInstances,
            double chordMm,
            double snapMm,
            int maxCandidates)
        {
            // Correct contour nesting boundary buffer:
            // expanded polygon (gap/2) must stay >= (margin + gap/2) from sheet edge
            double boundaryBufferMm = marginMm + gapMm / 2.0;

            double usableWmm = sheetWmm - 2 * boundaryBufferMm;
            double usableHmm = sheetHmm - 2 * boundaryBufferMm;
            if (usableWmm <= 0 || usableHmm <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            long usableW = ToInt(usableWmm);
            long usableH = ToInt(usableHmm);

            // Expand instances
            var instances = new List<PartDefinition>();
            foreach (var d in defs)
                for (int i = 0; i < d.Quantity; i++)
                    instances.Add(d);

            // Sort by contour area (fallback bbox area)
            instances.Sort((a, b) =>
            {
                long areaA = a.OuterArea2Abs > 0 ? a.OuterArea2Abs : ToInt(a.Width) * ToInt(a.Height);
                long areaB = b.OuterArea2Abs > 0 ? b.OuterArea2Abs : ToInt(b.Width) * ToInt(b.Height);
                return areaB.CompareTo(areaA);
            });

            // Precompute rotated + offset polygons per unique block (per rotation)
            // Key: block name + rotation
            var polyCache = new Dictionary<string, RotatedPoly>(StringComparer.OrdinalIgnoreCase);

            RotatedPoly GetRot(PartDefinition part, int rotDeg)
            {
                string key = part.BlockName + "||" + rotDeg.ToString(CultureInfo.InvariantCulture) + "||gap:" + gapMm.ToString("0.###", CultureInfo.InvariantCulture);
                if (polyCache.TryGetValue(key, out var rp))
                    return rp;

                // 1) base contour (ensure we have it; fallback rectangle)
                Path64 basePoly = part.OuterContour0;
                if (basePoly == null || basePoly.Count < 3)
                {
                    basePoly = MakeRectPolyScaled(part.MinX, part.MinY, part.MaxX, part.MaxY);
                }

                // 2) rotate
                Path64 rotPoly = RotatePoly(basePoly, rotDeg);

                // 3) offset outward by gap/2
                double delta = (gapMm / 2.0) * SCALE;
                Path64 offset = OffsetLargest(rotPoly, delta);

                // If offset fails, fallback to rotPoly (still works, just no gap safety)
                if (offset == null || offset.Count < 3)
                    offset = rotPoly;

                // Clean duplicates
                offset = CleanPath(offset);

                var bbox = GetBounds(offset);

                // anchors (4)
                var anchors = GetAnchors(offset);

                // area
                long area2Abs = Area2Abs(rotPoly);

                rp = new RotatedPoly
                {
                    RotDeg = rotDeg,
                    RotRad = rotDeg * Math.PI / 180.0,
                    PolyRot = rotPoly,
                    PolyOffset = offset,
                    OffsetBounds = bbox,
                    Anchors = anchors,
                    RotArea2Abs = area2Abs
                };

                polyCache[key] = rp;
                return rp;
            }

            var sheets = new List<SheetContourState>();

            SheetContourState NewSheet()
            {
                var s = new SheetContourState
                {
                    Index = sheets.Count + 1,
                    OriginXmm = baseOriginXmm + (sheets.Count) * (sheetWmm + 60.0),
                    OriginYmm = baseOriginYmm
                };

                DrawSheetOutline(s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, modelSpace, materialLabel, s.Index, NestingMode.ContourLevel1);
                sheets.Add(s);
                return s;
            }

            var cur = NewSheet();

            int placed = 0;

            foreach (var part in instances)
            {
                bool placedThis = false;

                // Try on existing sheets first
                for (int si = 0; si < sheets.Count; si++)
                {
                    if (TryPlaceContourOnSheet(
                        sheets[si],
                        part,
                        usableW,
                        usableH,
                        maxCandidates,
                        GetRot,
                        out var placement))
                    {
                        AddPlacedToDwg(
                            modelSpace,
                            part,
                            sheets[si],
                            boundaryBufferMm,
                            placement.InsertX,
                            placement.InsertY,
                            placement.RotRad);

                        // store translated offset poly for collision checks
                        sheets[si].Placed.Add(new PlacedContour
                        {
                            OffsetPoly = placement.OffsetPolyTranslated,
                            BBox = placement.OffsetBBoxTranslated
                        });

                        sheets[si].PlacedCount++;
                        sheets[si].UsedArea2Abs += placement.RotArea2Abs;

                        placed++;
                        progress.Step($"Placed {placed}/{totalInstances} (Contour L1)");

                        placedThis = true;
                        break;
                    }
                }

                if (placedThis)
                    continue;

                // Need new sheet
                cur = NewSheet();

                if (!TryPlaceContourOnSheet(cur, part, usableW, usableH, maxCandidates, GetRot, out var placement2))
                    throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

                AddPlacedToDwg(modelSpace, part, cur, boundaryBufferMm, placement2.InsertX, placement2.InsertY, placement2.RotRad);

                cur.Placed.Add(new PlacedContour
                {
                    OffsetPoly = placement2.OffsetPolyTranslated,
                    BBox = placement2.OffsetBBoxTranslated
                });

                cur.PlacedCount++;
                cur.UsedArea2Abs += placement2.RotArea2Abs;

                placed++;
                progress.Step($"Placed {placed}/{totalInstances} (Contour L1)");
            }

            // Optional sheet label with fill %
            long usableArea2 = usableW * usableH; // scaled^2
            foreach (var s in sheets)
            {
                double fill = usableArea2 > 0 ? (double)s.UsedArea2Abs / usableArea2 * 100.0 : 0.0;
                AddSheetLabel(modelSpace, s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, s.Index, fill);
            }

            return sheets.Count;
        }

        private struct ContourPlacement
        {
            public long InsertX;
            public long InsertY;
            public double RotRad;

            public Path64 OffsetPolyTranslated;
            public LongRect OffsetBBoxTranslated;

            public long RotArea2Abs;
        }

        private struct RotatedPoly
        {
            public int RotDeg;
            public double RotRad;

            public Path64 PolyRot;
            public Path64 PolyOffset;

            public LongRect OffsetBounds;

            public Point64[] Anchors; // 4 anchors from offset poly

            public long RotArea2Abs;
        }

        private static bool TryPlaceContourOnSheet(
            SheetContourState sheet,
            PartDefinition part,
            long usableW,
            long usableH,
            int maxCandidates,
            Func<PartDefinition, int, RotatedPoly> getRot,
            out ContourPlacement placement)
        {
            placement = default;

            // Try all 0/90/180/270
            // Strategy: search candidates in "bottom-left" order (min Y then min X),
            // accept first feasible.
            foreach (int rotDeg in RotationsDeg)
            {
                var rp = getRot(part, rotDeg);

                // Quick fit check using offset bbox size
                long offW = rp.OffsetBounds.MaxX - rp.OffsetBounds.MinX;
                long offH = rp.OffsetBounds.MaxY - rp.OffsetBounds.MinY;
                if (offW <= 0 || offH <= 0 || offW > usableW || offH > usableH)
                    continue;

                // Generate candidates
                var candidates = GenerateCandidates(sheet, rp, maxCandidates);

                // Sort by resulting bbox minY, then minX
                candidates.Sort((a, b) =>
                {
                    long aMinY = a.InsY + rp.OffsetBounds.MinY;
                    long bMinY = b.InsY + rp.OffsetBounds.MinY;
                    int cmp = aMinY.CompareTo(bMinY);
                    if (cmp != 0) return cmp;

                    long aMinX = a.InsX + rp.OffsetBounds.MinX;
                    long bMinX = b.InsX + rp.OffsetBounds.MinX;
                    return aMinX.CompareTo(bMinX);
                });

                foreach (var cand in candidates)
                {
                    long minX = cand.InsX + rp.OffsetBounds.MinX;
                    long minY = cand.InsY + rp.OffsetBounds.MinY;
                    long maxX = cand.InsX + rp.OffsetBounds.MaxX;
                    long maxY = cand.InsY + rp.OffsetBounds.MaxY;

                    if (minX < 0 || minY < 0 || maxX > usableW || maxY > usableH)
                        continue;

                    // Translate offset poly for collision checks
                    var moved = TranslatePath(rp.PolyOffset, cand.InsX, cand.InsY);
                    var movedBBox = new LongRect { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };

                    bool overlap = false;

                    foreach (var placed in sheet.Placed)
                    {
                        if (!RectsOverlap(movedBBox, placed.BBox))
                            continue;

                        if (PolygonsOverlapAreaPositive(moved, placed.OffsetPoly))
                        {
                            overlap = true;
                            break;
                        }
                    }

                    if (overlap)
                        continue;

                    placement = new ContourPlacement
                    {
                        InsertX = cand.InsX,
                        InsertY = cand.InsY,
                        RotRad = rp.RotRad,

                        OffsetPolyTranslated = moved,
                        OffsetBBoxTranslated = movedBBox,

                        RotArea2Abs = rp.RotArea2Abs
                    };
                    return true;
                }
            }

            return false;
        }

        private struct CandidateIns
        {
            public long InsX, InsY;
        }

        private static List<CandidateIns> GenerateCandidates(SheetContourState sheet, RotatedPoly rp, int maxCandidates)
        {
            var result = new List<CandidateIns>(Math.Min(maxCandidates, 2048));
            var seen = new HashSet<(long, long)>();

            void Add(long ix, long iy)
            {
                if (result.Count >= maxCandidates)
                    return;

                var key = (ix, iy);
                if (seen.Add(key))
                    result.Add(new CandidateIns { InsX = ix, InsY = iy });
            }

            // Base candidate: align bbox-min to (0,0)
            Add(-rp.OffsetBounds.MinX, -rp.OffsetBounds.MinY);

            // Bounding-box edge candidates (fast)
            // x candidates from right edges, y from top edges
            var xSet = new HashSet<long> { 0 };
            var ySet = new HashSet<long> { 0 };

            foreach (var p in sheet.Placed)
            {
                xSet.Add(p.BBox.MaxX);
                ySet.Add(p.BBox.MaxY);
            }

            var xs = xSet.OrderBy(v => v).Take(120).ToList();
            var ys = ySet.OrderBy(v => v).Take(120).ToList();

            foreach (var y in ys)
            {
                foreach (var x in xs)
                {
                    // desired bbox-min = (x,y)
                    Add(x - rp.OffsetBounds.MinX, y - rp.OffsetBounds.MinY);
                    if (result.Count >= maxCandidates) break;
                }
                if (result.Count >= maxCandidates) break;
            }

            // Vertex-based candidates (better interlocking)
            // Align moving anchors to sampled vertices of placed polygons.
            if (result.Count < maxCandidates && sheet.Placed.Count > 0)
            {
                foreach (var placed in sheet.Placed)
                {
                    int n = placed.OffsetPoly.Count;
                    if (n < 3) continue;

                    int step = Math.Max(1, n / 30); // sample ~30 points max

                    for (int i = 0; i < n; i += step)
                    {
                        var v = placed.OffsetPoly[i];

                        for (int a = 0; a < rp.Anchors.Length; a++)
                        {
                            var m = rp.Anchors[a];
                            Add(v.X - m.X, v.Y - m.Y);
                            if (result.Count >= maxCandidates) break;
                        }

                        if (result.Count >= maxCandidates) break;
                    }

                    if (result.Count >= maxCandidates) break;
                }
            }

            return result;
        }

        private static void AddPlacedToDwg(
            BlockRecord modelSpace,
            PartDefinition part,
            SheetContourState sheet,
            double boundaryBufferMm,
            long insertXScaled,
            long insertYScaled,
            double rotRad)
        {
            double insXmm = sheet.OriginXmm + boundaryBufferMm + (double)insertXScaled / SCALE;
            double insYmm = sheet.OriginYmm + boundaryBufferMm + (double)insertYScaled / SCALE;

            var ins = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insXmm, insYmm, 0.0),
                Rotation = rotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            };
            modelSpace.Entities.Add(ins);
        }

        // -----------------------------------------
        // Part loading + contour extraction
        // -----------------------------------------
        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc, LaserCutRunSettings settings)
        {
            if (doc == null)
                yield break;

            foreach (var br in doc.BlockRecords)
            {
                var block = br;
                if (block == null) continue;

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

                // material exact from block token written by combiner
                string material = "UNKNOWN";
                MaterialNameCodec.TryExtractFromBlockName(name, out material);
                material = MaterialNameCodec.Normalize(material);

                // bbox (mm)
                if (!TryGetBlockBbox(block, out double minX, out double minY, out double maxX, out double maxY))
                    continue;

                double w = maxX - minX;
                double h = maxY - minY;
                if (w <= 0 || h <= 0)
                    continue;

                // contour (only needed for contour mode; but we can compute always)
                Path64 contour = null;
                long area2Abs = 0;

                try
                {
                    contour = ExtractOuterContourScaled(block, chordMm: settings.ContourChordMm, snapMm: settings.ContourSnapMm);
                    contour = CleanPath(contour);
                    area2Abs = Area2Abs(contour);
                }
                catch
                {
                    contour = null;
                    area2Abs = 0;
                }

                // fallback rectangle if contour failed
                if (contour == null || contour.Count < 3)
                {
                    contour = MakeRectPolyScaled(minX, minY, maxX, maxY);
                    area2Abs = Area2Abs(contour);
                }

                yield return new PartDefinition
                {
                    Block = block,
                    BlockName = name,
                    Quantity = Math.Max(1, qty),
                    MaterialExact = material,

                    MinX = minX,
                    MinY = minY,
                    MaxX = maxX,
                    MaxY = maxY,
                    Width = w,
                    Height = h,

                    OuterContour0 = contour,
                    OuterArea2Abs = area2Abs
                };
            }
        }

        private static bool TryGetBlockBbox(BlockRecord block, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            bool any = false;

            foreach (var ent in block.Entities)
            {
                try
                {
                    var bb = ent.GetBoundingBox();
                    var a = bb.Min;
                    var b = bb.Max;

                    if (a.X < minX) minX = a.X;
                    if (a.Y < minY) minY = a.Y;
                    if (b.X > maxX) maxX = b.X;
                    if (b.Y > maxY) maxY = b.Y;

                    any = true;
                }
                catch
                {
                    // ignore entities without bbox
                }
            }

            if (!any || minX == double.MaxValue || maxX == double.MinValue)
                return false;

            return true;
        }

        private static Path64 ExtractOuterContourScaled(BlockRecord block, double chordMm, double snapMm)
        {
            if (block == null)
                return null;

            chordMm = Math.Max(0.10, chordMm);
            snapMm = Math.Max(0.01, snapMm);

            // 1) collect segments (scaled)
            var segs = new List<(Point64 A, Point64 B)>();

            foreach (var ent in block.Entities)
            {
                if (ent == null) continue;

                if (ent is Line ln)
                {
                    segs.Add((
                        Snap(ToP64(ln.StartPoint), snapMm),
                        Snap(ToP64(ln.EndPoint), snapMm)));
                }
                else if (ent is Arc arc)
                {
                    AddArcSegments(segs, arc.Center, arc.Radius, arc.StartAngle, arc.EndAngle, chordMm, snapMm);
                }
                else if (ent is Circle cir)
                {
                    AddCircleSegments(segs, cir.Center, cir.Radius, chordMm, snapMm);
                }
                else
                {
                    // polylines (reflection-safe)
                    if (TryAddPolylineSegments(ent, segs, chordMm, snapMm))
                        continue;
                }
            }

            // Not enough data
            if (segs.Count < 3)
                return null;

            // 2) build loops from segments
            var loops = BuildClosedLoops(segs);

            if (loops.Count > 0)
            {
                // pick largest abs area
                Path64 best = null;
                long bestArea = 0;

                foreach (var loop in loops)
                {
                    long a2 = Area2Abs(loop);
                    if (a2 > bestArea)
                    {
                        bestArea = a2;
                        best = loop;
                    }
                }

                return best;
            }

            // 3) fallback: convex hull of all points
            var pts = new List<Point64>(segs.Count * 2);
            foreach (var s in segs)
            {
                pts.Add(s.A);
                pts.Add(s.B);
            }

            var hull = ConvexHull(pts);
            return hull;
        }

        private static List<Path64> BuildClosedLoops(List<(Point64 A, Point64 B)> segs)
        {
            var loops = new List<Path64>();
            if (segs == null || segs.Count == 0)
                return loops;

            // adjacency map
            var adj = new Dictionary<(long, long), List<int>>();
            var used = new bool[segs.Count];

            (long, long) Key(Point64 p) => (p.X, p.Y);

            for (int i = 0; i < segs.Count; i++)
            {
                var s = segs[i];
                var kA = Key(s.A);
                var kB = Key(s.B);

                if (!adj.TryGetValue(kA, out var la)) { la = new List<int>(); adj[kA] = la; }
                la.Add(i);

                if (!adj.TryGetValue(kB, out var lb)) { lb = new List<int>(); adj[kB] = lb; }
                lb.Add(i);
            }

            for (int i = 0; i < segs.Count; i++)
            {
                if (used[i]) continue;

                var s0 = segs[i];
                var start = s0.A;
                var startK = Key(start);

                var path = new Path64();
                path.Add(start);

                // walk
                Point64 cur = s0.B;
                var curK = Key(cur);

                used[i] = true;
                path.Add(cur);

                var prevK = startK;

                int guard = 0;
                while (curK != startK && guard++ < segs.Count + 10)
                {
                    if (!adj.TryGetValue(curK, out var incident))
                        break;

                    int nextSeg = -1;

                    // choose an unused segment, prefer not going back to prev point
                    foreach (int si in incident)
                    {
                        if (used[si]) continue;
                        var s = segs[si];
                        var aK = Key(s.A);
                        var bK = Key(s.B);

                        var otherK = (aK == curK) ? bK : (bK == curK ? aK : curK);

                        if (otherK != prevK)
                        {
                            nextSeg = si;
                            break;
                        }

                        if (nextSeg < 0)
                            nextSeg = si;
                    }

                    if (nextSeg < 0)
                        break;

                    used[nextSeg] = true;

                    var ns = segs[nextSeg];
                    var aK2 = Key(ns.A);
                    var bK2 = Key(ns.B);

                    Point64 nextPt;
                    (long, long) nextK;

                    if (aK2 == curK)
                    {
                        nextPt = ns.B;
                        nextK = bK2;
                    }
                    else
                    {
                        nextPt = ns.A;
                        nextK = aK2;
                    }

                    // avoid duplicates
                    if (path.Count == 0 || path[path.Count - 1].X != nextPt.X || path[path.Count - 1].Y != nextPt.Y)
                        path.Add(nextPt);

                    prevK = curK;
                    curK = nextK;
                    cur = nextPt;
                }

                // closed?
                if (curK == startK && path.Count >= 4)
                {
                    // remove last duplicate if it equals first
                    if (path.Count > 1 && path[path.Count - 1].X == path[0].X && path[path.Count - 1].Y == path[0].Y)
                        path.RemoveAt(path.Count - 1);

                    path = CleanPath(path);

                    if (path != null && path.Count >= 3)
                        loops.Add(path);
                }
            }

            return loops;
        }

        private static bool TryAddPolylineSegments(Entity ent, List<(Point64 A, Point64 B)> segs, double chordMm, double snapMm)
        {
            // Supports both LwPolyline and Polyline (reflection-safe for vertex access + bulge)
            try
            {
                var t = ent.GetType();
                string tn = t.Name ?? "";

                if (!tn.Contains("Polyline", StringComparison.OrdinalIgnoreCase))
                    return false;

                var vertsProp = t.GetProperty("Vertices");
                var vertsObj = vertsProp?.GetValue(ent);
                if (vertsObj == null)
                    return false;

                var vertsEnum = vertsObj as System.Collections.IEnumerable;
                if (vertsEnum == null)
                    return false;

                var verts = new List<(double X, double Y, double Bulge)>();
                foreach (var v in vertsEnum)
                {
                    if (TryGetVertexXYB(v, out double x, out double y, out double b))
                        verts.Add((x, y, b));
                }

                if (verts.Count < 2)
                    return false;

                bool closed = false;
                var closedProp = t.GetProperty("IsClosed") ?? t.GetProperty("Closed");
                if (closedProp != null && closedProp.PropertyType == typeof(bool))
                    closed = (bool)closedProp.GetValue(ent);

                int count = verts.Count;
                int last = closed ? count : count - 1;

                for (int i = 0; i < last; i++)
                {
                    var v1 = verts[i];
                    var v2 = verts[(i + 1) % count];

                    if (Math.Abs(v1.Bulge) < 1e-12)
                    {
                        segs.Add((
                            Snap(ToP64(v1.X, v1.Y), snapMm),
                            Snap(ToP64(v2.X, v2.Y), snapMm)));
                    }
                    else
                    {
                        AddBulgeArcSegments(segs, v1, v2, chordMm, snapMm);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetVertexXYB(object v, out double x, out double y, out double bulge)
        {
            x = y = 0.0;
            bulge = 0.0;

            if (v == null) return false;

            try
            {
                var t = v.GetType();

                var pb = t.GetProperty("Bulge");
                if (pb != null)
                {
                    object bv = pb.GetValue(v);
                    if (bv is double bd) bulge = bd;
                }

                // Common patterns:
                // - X / Y properties
                // - Location (XYZ)
                // - Point (XYZ)
                var px = t.GetProperty("X");
                var py = t.GetProperty("Y");

                if (px != null && py != null)
                {
                    x = Convert.ToDouble(px.GetValue(v), CultureInfo.InvariantCulture);
                    y = Convert.ToDouble(py.GetValue(v), CultureInfo.InvariantCulture);
                    return true;
                }

                var ploc = t.GetProperty("Location") ?? t.GetProperty("Point");
                if (ploc != null)
                {
                    var loc = ploc.GetValue(v);
                    if (loc != null)
                    {
                        var lt = loc.GetType();
                        var lx = lt.GetProperty("X");
                        var ly = lt.GetProperty("Y");
                        if (lx != null && ly != null)
                        {
                            x = Convert.ToDouble(lx.GetValue(loc), CultureInfo.InvariantCulture);
                            y = Convert.ToDouble(ly.GetValue(loc), CultureInfo.InvariantCulture);
                            return true;
                        }
                    }
                }
            }
            catch { }

            return false;
        }

        private static void AddBulgeArcSegments(List<(Point64 A, Point64 B)> segs, (double X, double Y, double Bulge) v1, (double X, double Y, double Bulge) v2, double chordMm, double snapMm)
        {
            double b = v1.Bulge;
            if (Math.Abs(b) < 1e-12)
            {
                segs.Add((Snap(ToP64(v1.X, v1.Y), snapMm), Snap(ToP64(v2.X, v2.Y), snapMm)));
                return;
            }

            double x1 = v1.X, y1 = v1.Y;
            double x2 = v2.X, y2 = v2.Y;

            double dx = x2 - x1;
            double dy = y2 - y1;
            double L = Math.Sqrt(dx * dx + dy * dy);
            if (L < 1e-9)
                return;

            double theta = 4.0 * Math.Atan(b); // signed
            double sinHalf = Math.Sin(theta / 2.0);
            if (Math.Abs(sinHalf) < 1e-12)
            {
                segs.Add((Snap(ToP64(x1, y1), snapMm), Snap(ToP64(x2, y2), snapMm)));
                return;
            }

            double R = L / (2.0 * sinHalf);
            double Rabs = Math.Abs(R);

            double mx = (x1 + x2) / 2.0;
            double my = (y1 + y2) / 2.0;

            double d = Math.Sqrt(Math.Max(0.0, Rabs * Rabs - (L / 2.0) * (L / 2.0)));

            // left normal
            double nx = -dy / L;
            double ny = dx / L;

            // bulge sign controls center side
            double sign = b >= 0 ? 1.0 : -1.0;

            double cx = mx + sign * nx * d;
            double cy = my + sign * ny * d;

            double a1 = Math.Atan2(y1 - cy, x1 - cx);
            // sweep = theta
            int segCount = Math.Max(8, (int)Math.Ceiling((Rabs * Math.Abs(theta)) / Math.Max(0.10, chordMm)));
            segCount = Math.Min(segCount, 720);

            double step = theta / segCount;

            Point64 prev = Snap(ToP64(x1, y1), snapMm);
            for (int i = 1; i <= segCount; i++)
            {
                double ang = a1 + step * i;
                double px = cx + Rabs * Math.Cos(ang);
                double py = cy + Rabs * Math.Sin(ang);

                var cur = Snap(ToP64(px, py), snapMm);
                segs.Add((prev, cur));
                prev = cur;
            }
        }

        private static void AddArcSegments(
            List<(Point64 A, Point64 B)> segs,
            XYZ center,
            double radius,
            double startAngle,
            double endAngle,
            double chordMm,
            double snapMm)
        {
            // DXF arc angles are generally degrees; SolidWorks exports usually degrees.
            // If your file is radians, you can switch here.
            double sa = DegreesToRadiansIfNeeded(startAngle);
            double ea = DegreesToRadiansIfNeeded(endAngle);

            // CCW sweep
            double sweep = ea - sa;
            while (sweep < 0) sweep += 2.0 * Math.PI;
            if (sweep <= 1e-12) sweep = 2.0 * Math.PI;

            double r = Math.Abs(radius);
            if (r <= 1e-9) return;

            int segCount = Math.Max(8, (int)Math.Ceiling((r * sweep) / Math.Max(0.10, chordMm)));
            segCount = Math.Min(segCount, 1440);

            Point64 prev = Snap(ToP64(center.X + r * Math.Cos(sa), center.Y + r * Math.Sin(sa)), snapMm);

            for (int i = 1; i <= segCount; i++)
            {
                double ang = sa + sweep * i / segCount;
                var cur = Snap(ToP64(center.X + r * Math.Cos(ang), center.Y + r * Math.Sin(ang)), snapMm);
                segs.Add((prev, cur));
                prev = cur;
            }
        }

        private static void AddCircleSegments(List<(Point64 A, Point64 B)> segs, XYZ center, double radius, double chordMm, double snapMm)
        {
            double r = Math.Abs(radius);
            if (r <= 1e-9) return;

            double sweep = 2.0 * Math.PI;
            int segCount = Math.Max(16, (int)Math.Ceiling((r * sweep) / Math.Max(0.10, chordMm)));
            segCount = Math.Min(segCount, 2880);

            Point64 first = Snap(ToP64(center.X + r, center.Y), snapMm);
            Point64 prev = first;

            for (int i = 1; i <= segCount; i++)
            {
                double ang = sweep * i / segCount;
                var cur = Snap(ToP64(center.X + r * Math.Cos(ang), center.Y + r * Math.Sin(ang)), snapMm);
                segs.Add((prev, cur));
                prev = cur;
            }
        }

        private static double DegreesToRadiansIfNeeded(double angle)
        {
            // If value is > 2π by a lot, assume degrees.
            if (Math.Abs(angle) > 10.0) // 10 rad ~ 572°
                return angle * Math.PI / 180.0;

            // Many SolidWorks exports have 0/90/180/270 => >10 triggers degrees, so OK.
            // Small angles could be ambiguous; leave as-is.
            return angle;
        }

        // -----------------------------------------
        // Polygon helpers (Clipper units)
        // -----------------------------------------
        private static long ToInt(double mm) => (long)Math.Round(mm * SCALE);

        private static Point64 ToP64(XYZ p) => new Point64(ToInt(p.X), ToInt(p.Y));
        private static Point64 ToP64(double x, double y) => new Point64(ToInt(x), ToInt(y));

        private static Point64 Snap(Point64 p, double snapMm)
        {
            long grid = Math.Max(1, (long)Math.Round(snapMm * SCALE));
            long sx = (long)Math.Round((double)p.X / grid) * grid;
            long sy = (long)Math.Round((double)p.Y / grid) * grid;
            return new Point64(sx, sy);
        }

        private static Path64 CleanPath(Path64 path)
        {
            if (path == null || path.Count == 0)
                return path;

            var res = new Path64();
            Point64 prev = path[0];
            res.Add(prev);

            for (int i = 1; i < path.Count; i++)
            {
                var cur = path[i];
                if (cur.X == prev.X && cur.Y == prev.Y)
                    continue;

                res.Add(cur);
                prev = cur;
            }

            // If last == first, drop last
            if (res.Count > 1 && res[0].X == res[res.Count - 1].X && res[0].Y == res[res.Count - 1].Y)
                res.RemoveAt(res.Count - 1);

            return res.Count >= 3 ? res : res;
        }

        private static Path64 MakeRectPolyScaled(double minX, double minY, double maxX, double maxY)
        {
            long x1 = ToInt(minX);
            long y1 = ToInt(minY);
            long x2 = ToInt(maxX);
            long y2 = ToInt(maxY);

            var p = new Path64
            {
                new Point64(x1, y1),
                new Point64(x2, y1),
                new Point64(x2, y2),
                new Point64(x1, y2)
            };
            return p;
        }

        private static Path64 RotatePoly(Path64 p, int rotDeg)
        {
            if (p == null) return null;
            rotDeg = ((rotDeg % 360) + 360) % 360;

            var r = new Path64(p.Count);

            foreach (var pt in p)
            {
                long x = pt.X;
                long y = pt.Y;

                switch (rotDeg)
                {
                    case 0:
                        r.Add(new Point64(x, y));
                        break;
                    case 90:
                        r.Add(new Point64(-y, x));
                        break;
                    case 180:
                        r.Add(new Point64(-x, -y));
                        break;
                    case 270:
                        r.Add(new Point64(y, -x));
                        break;
                    default:
                        // should never happen in our rotations list
                        double rad = rotDeg * Math.PI / 180.0;
                        long xr = (long)Math.Round(x * Math.Cos(rad) - y * Math.Sin(rad));
                        long yr = (long)Math.Round(x * Math.Sin(rad) + y * Math.Cos(rad));
                        r.Add(new Point64(xr, yr));
                        break;
                }
            }

            return r;
        }

        private static Path64 OffsetLargest(Path64 poly, double delta)
        {
            if (poly == null || poly.Count < 3)
                return null;

            var co = new ClipperOffset();
            // Round is safer for offsets (reduces spikes on acute angles)
            co.AddPath(poly, JoinType.Round, EndType.Polygon);

            var sol = new Paths64();
            co.Execute(delta, sol);

            if (sol == null || sol.Count == 0)
                return null;

            // pick largest abs area
            Path64 best = null;
            long bestArea = 0;

            foreach (var p in sol)
            {
                long a2 = Area2Abs(p);
                if (a2 > bestArea)
                {
                    bestArea = a2;
                    best = p;
                }
            }

            return best;
        }

        private static LongRect GetBounds(Path64 p)
        {
            long minX = long.MaxValue, minY = long.MaxValue;
            long maxX = long.MinValue, maxY = long.MinValue;

            foreach (var pt in p)
            {
                if (pt.X < minX) minX = pt.X;
                if (pt.Y < minY) minY = pt.Y;
                if (pt.X > maxX) maxX = pt.X;
                if (pt.Y > maxY) maxY = pt.Y;
            }

            return new LongRect { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };
        }

        private static Point64[] GetAnchors(Path64 p)
        {
            // anchors based on extremes
            Point64 bl = p[0], br = p[0], tl = p[0], tr = p[0];

            foreach (var pt in p)
            {
                if (pt.Y < bl.Y || (pt.Y == bl.Y && pt.X < bl.X)) bl = pt;
                if (pt.Y < br.Y || (pt.Y == br.Y && pt.X > br.X)) br = pt;
                if (pt.Y > tl.Y || (pt.Y == tl.Y && pt.X < tl.X)) tl = pt;
                if (pt.Y > tr.Y || (pt.Y == tr.Y && pt.X > tr.X)) tr = pt;
            }

            return new[] { bl, br, tl, tr };
        }

        private static Path64 TranslatePath(Path64 p, long dx, long dy)
        {
            var r = new Path64(p.Count);
            foreach (var pt in p)
                r.Add(new Point64(pt.X + dx, pt.Y + dy));
            return r;
        }

        private static bool RectsOverlap(LongRect a, LongRect b)
        {
            return !(a.MaxX <= b.MinX || b.MaxX <= a.MinX || a.MaxY <= b.MinY || b.MaxY <= a.MinY);
        }

        private static bool PolygonsOverlapAreaPositive(Path64 a, Path64 b)
        {
            // Intersection area > 0 => overlap
            var clipper = new Clipper64();
            clipper.AddSubject(a);
            clipper.AddClip(b);

            var sol = new Paths64();
            clipper.Execute(ClipperClipType.Intersection, FillRule.NonZero, sol);

            if (sol == null || sol.Count == 0)
                return false;

            foreach (var p in sol)
            {
                if (Area2Abs(p) > 0)
                    return true;
            }

            return false;
        }

        private static long Area2Abs(Path64 p)
        {
            if (p == null || p.Count < 3)
                return 0;

            long sum = 0;
            int n = p.Count;

            for (int i = 0; i < n; i++)
            {
                var a = p[i];
                var b = p[(i + 1) % n];

                // cross
                sum += a.X * b.Y - b.X * a.Y;
            }

            return Math.Abs(sum);
        }

        private static Path64 ConvexHull(List<Point64> pts)
        {
            if (pts == null)
                return null;

            // remove duplicates
            var uniq = pts.Distinct().ToList();
            if (uniq.Count < 3)
                return null;

            uniq.Sort((a, b) =>
            {
                int c = a.X.CompareTo(b.X);
                if (c != 0) return c;
                return a.Y.CompareTo(b.Y);
            });

            long Cross(Point64 o, Point64 a, Point64 b)
                => (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X);

            var lower = new List<Point64>();
            foreach (var p in uniq)
            {
                while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], p) <= 0)
                    lower.RemoveAt(lower.Count - 1);
                lower.Add(p);
            }

            var upper = new List<Point64>();
            for (int i = uniq.Count - 1; i >= 0; i--)
            {
                var p = uniq[i];
                while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], p) <= 0)
                    upper.RemoveAt(upper.Count - 1);
                upper.Add(p);
            }

            lower.RemoveAt(lower.Count - 1);
            upper.RemoveAt(upper.Count - 1);

            var hull = new Path64();
            hull.AddRange(lower);
            hull.AddRange(upper);

            return hull.Count >= 3 ? hull : null;
        }

        // -----------------------------------------
        // Grouping (material exact)
        // -----------------------------------------
        private static Dictionary<string, string> BuildGroups(List<PartDefinition> defs, LaserCutRunSettings settings)
        {
            var groups = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!settings.SeparateByMaterialExact || !settings.OutputOneDwgPerMaterial)
            {
                groups["ALL"] = "ALL";
                return groups;
            }

            foreach (var d in defs)
            {
                string key = NormalizeKey(d.MaterialExact);
                if (!groups.ContainsKey(key))
                    groups[key] = d.MaterialExact;
            }

            if (groups.Count == 0)
                groups["UNKNOWN"] = "UNKNOWN";

            return groups;
        }

        private static string NormalizeKey(string s)
        {
            s = (s ?? "").Trim();
            return string.IsNullOrWhiteSpace(s) ? "UNKNOWN" : s.ToUpperInvariant();
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

        // -----------------------------------------
        // Option 2 preview filtering
        // -----------------------------------------
        private static void FilterSourcePreviewToTheseBlocks(CadDocument doc, HashSet<string> keepBlockNames)
        {
            if (doc == null || keepBlockNames == null || keepBlockNames.Count == 0)
                return;

            BlockRecord modelSpace;
            try { modelSpace = doc.BlockRecords["*Model_Space"]; }
            catch { return; }

            var inserts = modelSpace.Entities.OfType<Insert>().ToList();

            var keepRanges = new List<(double minX, double maxX)>();

            // Build lookup bbox per block
            var defMap = new Dictionary<string, (double minX, double minY, double maxX, double maxY)>(StringComparer.OrdinalIgnoreCase);
            foreach (var br in doc.BlockRecords)
            {
                if (br == null) continue;
                string n = br.Name ?? "";
                if (!n.StartsWith("P_", StringComparison.OrdinalIgnoreCase)) continue;

                if (TryGetBlockBbox(br, out double mnX, out double mnY, out double mxX, out double mxY))
                    defMap[n] = (mnX, mnY, mxX, mxY);
            }

            foreach (var ins in inserts)
            {
                var blk = ins.Block;
                if (blk == null) continue;

                string bn = blk.Name ?? "";
                if (!bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!keepBlockNames.Contains(bn))
                    continue;

                if (!defMap.TryGetValue(bn, out var bb))
                    continue;

                double ix = ins.InsertPoint.X;
                double minX = ix + bb.minX;
                double maxX = ix + bb.maxX;
                if (minX > maxX) { var t = minX; minX = maxX; maxX = t; }
                keepRanges.Add((minX, maxX));
            }

            const double pad = 120.0;

            bool IsNear(double x)
            {
                foreach (var r in keepRanges)
                {
                    if (x >= r.minX - pad && x <= r.maxX + pad)
                        return true;
                }
                return false;
            }

            var remove = new List<Entity>();

            foreach (var e in modelSpace.Entities)
            {
                if (e is Insert ins)
                {
                    var blk = ins.Block;
                    if (blk == null) continue;

                    string bn = blk.Name ?? "";
                    if (bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase) && !keepBlockNames.Contains(bn))
                        remove.Add(e);
                }
                else if (e is MText mt)
                {
                    if (!IsNear(mt.InsertPoint.X))
                        remove.Add(e);
                }
                else
                {
                    // keep other entities (axes, etc.) to avoid removing useful stuff
                }
            }

            foreach (var e in remove.Distinct())
            {
                try { modelSpace.Entities.Remove(e); } catch { }
            }
        }

        private static void GetModelSpaceExtents(CadDocument doc, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            BlockRecord modelSpace;
            try { modelSpace = doc.BlockRecords["*Model_Space"]; }
            catch
            {
                minX = minY = maxX = maxY = 0.0;
                return;
            }

            bool any = false;

            foreach (var ent in modelSpace.Entities)
            {
                try
                {
                    var bb = ent.GetBoundingBox();
                    var a = bb.Min;
                    var b = bb.Max;

                    if (a.X < minX) minX = a.X;
                    if (a.Y < minY) minY = a.Y;
                    if (b.X > maxX) maxX = b.X;
                    if (b.Y > maxY) maxY = b.Y;

                    any = true;
                }
                catch { }
            }

            if (!any || minX == double.MaxValue || maxX == double.MinValue)
                minX = minY = maxX = maxY = 0.0;
        }

        // -----------------------------------------
        // DWG visuals (sheet outline + labels)
        // -----------------------------------------
        private static void DrawSheetOutline(
            double originXmm,
            double originYmm,
            double sheetWmm,
            double sheetHmm,
            BlockRecord modelSpace,
            string materialLabel,
            int sheetIndex,
            NestingMode mode)
        {
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm, originYmm, 0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm, 0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0), EndPoint = new XYZ(originXmm, originYmm + sheetHmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm, originYmm + sheetHmm, 0), EndPoint = new XYZ(originXmm, originYmm, 0) });

            string title =
                $"Sheet {sheetIndex}" +
                (string.IsNullOrWhiteSpace(materialLabel) || materialLabel.Equals("ALL", StringComparison.OrdinalIgnoreCase) ? "" : $" | {materialLabel}") +
                $" | {mode}";

            modelSpace.Entities.Add(new MText
            {
                Value = title,
                InsertPoint = new XYZ(originXmm + 10.0, originYmm + sheetHmm + 18.0, 0.0),
                Height = 20.0
            });
        }

        private static void AddSheetLabel(BlockRecord modelSpace, double originXmm, double originYmm, double sheetWmm, double sheetHmm, int sheetIndex, double fillPercent)
        {
            modelSpace.Entities.Add(new MText
            {
                Value = $"Fill: {fillPercent:0.0}%",
                InsertPoint = new XYZ(originXmm + sheetWmm - 220.0, originYmm + sheetHmm + 18.0, 0.0),
                Height = 18.0
            });
        }

        // -----------------------------------------
        // Log
        // -----------------------------------------
        private static void AppendNestLog(
            string logPath,
            string thicknessFile,
            string material,
            double sheetW,
            double sheetH,
            double thicknessMm,
            double gapMm,
            double marginMm,
            int sheets,
            int parts,
            string outDwg,
            NestingMode mode)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Nest run:");
                sb.AppendLine("  Thickness file: " + Path.GetFileName(thicknessFile));
                sb.AppendLine("  Material: " + material);
                sb.AppendLine($"  Mode: {mode}");
                sb.AppendLine($"  Sheet: {sheetW:0.###} x {sheetH:0.###} mm");
                sb.AppendLine($"  Thickness(mm): {thicknessMm:0.###}");
                sb.AppendLine($"  Gap(mm): {gapMm:0.###}  (auto >= thickness)");
                sb.AppendLine($"  Margin(mm): {marginMm:0.###}");
                sb.AppendLine($"  Sheets used: {sheets}");
                sb.AppendLine($"  Total parts: {parts}");
                sb.AppendLine("  Output: " + Path.GetFileName(outDwg));
                sb.AppendLine(new string('-', 70));

                File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // ignore logging failures
            }
        }
    }
}
