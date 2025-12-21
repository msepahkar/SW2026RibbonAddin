// Commands\dwg\LaserCutButton.cs
// DROP-IN: replace the entire file with this one.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;

// Avoid WinForms Point/Size ambiguity if you later add ACadSharp Point types
using WinPoint = System.Drawing.Point;
using WinSize = System.Drawing.Size;
using WinContentAlignment = System.Drawing.ContentAlignment;

// Avoid ambiguity with any "Application" symbol
using WinFormsApp = System.Windows.Forms.Application;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nCut";
        public string Tooltip => "Batch-nest thickness_*.dwg into sheets (0/90/180/270). Optional: separate by material + one DWG per material.";
        public string Hint => "Laser cut nesting";

        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            // Pick any thickness_*.dwg -> we will batch process all thickness_*.dwg in the same folder
            string anyThicknessDwg = SelectAnyThicknessDwg();
            if (string.IsNullOrEmpty(anyThicknessDwg))
                return;

            string folder = Path.GetDirectoryName(anyThicknessDwg) ?? "";
            if (!Directory.Exists(folder))
            {
                MessageBox.Show("Folder does not exist:\r\n" + folder, "Laser Cut", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var inputs = Directory.GetFiles(folder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
                .Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f) ?? "";
                    // exclude nested outputs and any accidental re-processing
                    return n.IndexOf("_nested", StringComparison.OrdinalIgnoreCase) < 0
                        && n.IndexOf("_nest_log", StringComparison.OrdinalIgnoreCase) < 0;
                })
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (inputs.Count == 0)
            {
                MessageBox.Show(
                    "No input files found.\r\n\r\nExpected files like:\r\n  thickness_2_5.dwg\r\n\r\nIn folder:\r\n" + folder,
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            LaserCutRunSettings settings;
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
            }

            // Pre-scan to compute total placements (for one global progress bar)
            int totalOverall = 0;
            foreach (var f in inputs)
            {
                try { totalOverall += Math.Max(0, DwgLaserNester.CountTotalInstances(f)); }
                catch { }
            }

            if (totalOverall <= 0)
            {
                MessageBox.Show(
                    "No plate blocks found in any thickness DWG.\r\n\r\nMake sure you ran Combine DWG and the thickness_*.dwg files contain P_* blocks.",
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            int placedOverall = 0;

            var summary = new StringBuilder();
            summary.AppendLine("BATCH LASER NEST SUMMARY");
            summary.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            summary.AppendLine($"Folder: {folder}");
            summary.AppendLine($"Files: {inputs.Count}");
            summary.AppendLine();
            summary.AppendLine($"SeparateByMaterial: {settings.SeparateByMaterial}");
            summary.AppendLine($"OutputOneDwgPerMaterial: {settings.SeparateByMaterial && settings.OutputOneDwgPerMaterial}");
            summary.AppendLine($"UsePerMaterialSheetPresets: {settings.UsePerMaterialSheetPresets}");
            summary.AppendLine($"GlobalSheet: {settings.DefaultSheet}");
            if (settings.SeparateByMaterial && settings.UsePerMaterialSheetPresets)
            {
                summary.AppendLine($"SteelSheet: {settings.SteelSheet}");
                summary.AppendLine($"AluminumSheet: {settings.AluminumSheet}");
                summary.AppendLine($"StainlessSheet: {settings.StainlessSheet}");
                summary.AppendLine($"OtherSheet: {settings.OtherSheet}");
            }
            summary.AppendLine();
            summary.AppendLine("Rotations: 0/90/180/270 (fixed)");
            summary.AppendLine("Gap+Margin: auto (>= thickness) + 10mm base margin");
            summary.AppendLine();
            summary.AppendLine("------------------------------------------------------------");
            summary.AppendLine();

            var errors = new List<string>();

            using (var progress = new LaserCutProgressForm(totalOverall))
            {
                progress.Show();
                WinFormsApp.DoEvents();

                foreach (var input in inputs)
                {
                    string fileName = Path.GetFileName(input) ?? input;

                    try
                    {
                        progress.SetStatus("Processing: " + fileName);

                        var res = DwgLaserNester.NestOneFile(
                            sourceDwgPath: input,
                            settings: settings,
                            progress: progress,
                            placedOverallRef: ref placedOverall,
                            totalOverall: totalOverall);

                        summary.AppendLine($"SOURCE: {fileName}");
                        summary.AppendLine($"  Blocks found/skipped: {res.CandidateBlocks}/{res.SkippedBlocks}");
                        foreach (var o in res.Outputs)
                        {
                            summary.AppendLine($"  [{o.MaterialType}] Sheets={o.SheetsUsed} Parts={o.TotalParts} Sheet={o.Sheet}");
                            summary.AppendLine($"     {Path.GetFileName(o.OutputDwgPath)}");
                        }

                        if (!string.IsNullOrEmpty(res.LogPath))
                            summary.AppendLine($"  Log: {Path.GetFileName(res.LogPath)}");

                        summary.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        string msg = $"FAILED: {fileName} -> {ex.Message}";
                        errors.Add(msg);

                        summary.AppendLine($"SOURCE: {fileName}");
                        summary.AppendLine("  ERROR: " + ex.Message);
                        summary.AppendLine();
                    }
                }

                progress.SetStatus("Writing summary...");
                WinFormsApp.DoEvents();
            }

            try
            {
                File.WriteAllText(Path.Combine(folder, "batch_nest_summary.txt"), summary.ToString(), Encoding.UTF8);
            }
            catch { }

            if (errors.Count == 0)
            {
                MessageBox.Show(
                    "Batch nesting finished.\r\n\r\n" +
                    "Summary:\r\n" + Path.Combine(folder, "batch_nest_summary.txt"),
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    "Batch nesting finished with errors.\r\n\r\n" +
                    string.Join("\r\n", errors.Take(12)) +
                    (errors.Count > 12 ? "\r\n..." : "") +
                    "\r\n\r\nSummary:\r\n" + Path.Combine(folder, "batch_nest_summary.txt"),
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        private static string SelectAnyThicknessDwg()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select ANY thickness_*.dwg (batch will process all thickness files in that folder)";
                dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.Multiselect = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.FileName;
            }
        }
    }

    internal struct SheetSize
    {
        public double W;
        public double H;

        public SheetSize(double w, double h)
        {
            W = w;
            H = h;
        }

        public override string ToString()
            => $"{W.ToString("0.###", CultureInfo.InvariantCulture)} x {H.ToString("0.###", CultureInfo.InvariantCulture)} mm";
    }

    internal sealed class LaserCutRunSettings
    {
        public bool SeparateByMaterial;
        public bool OutputOneDwgPerMaterial;
        public bool UsePerMaterialSheetPresets;

        public SheetSize DefaultSheet;

        public SheetSize SteelSheet;
        public SheetSize AluminumSheet;
        public SheetSize StainlessSheet;
        public SheetSize OtherSheet;

        public SheetSize GetSheetForMaterialType(string materialType)
        {
            if (!UsePerMaterialSheetPresets)
                return DefaultSheet;

            materialType = (materialType ?? "").Trim().ToUpperInvariant();

            if (materialType == MaterialTypeNormalizer.TYPE_STEEL) return SteelSheet;
            if (materialType == MaterialTypeNormalizer.TYPE_ALUMINUM) return AluminumSheet;
            if (materialType == MaterialTypeNormalizer.TYPE_STAINLESS) return StainlessSheet;

            return OtherSheet;
        }
    }

    internal static class DwgLaserNester
    {
        // Rotations are fixed: 0/90/180/270
        private static readonly int[] AllowedAnglesDeg = new[] { 0, 90, 180, 270 };

        internal sealed class MaterialNestResult
        {
            public string MaterialType;
            public SheetSize Sheet;
            public string OutputDwgPath;
            public int SheetsUsed;
            public int TotalParts;
        }

        internal sealed class NestRunResult
        {
            public string SourceDwgPath;
            public LaserCutRunSettings Settings;

            public List<MaterialNestResult> Outputs = new List<MaterialNestResult>();

            public int CandidateBlocks;
            public int SkippedBlocks;

            public string LogPath;
            public int LogLines;
        }

        private sealed class NestLog
        {
            private readonly List<string> _lines = new List<string>();
            public int Count => _lines.Count;

            public void Info(string msg)
            {
                if (string.IsNullOrWhiteSpace(msg)) return;
                _lines.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {msg}");
            }

            public void Warn(string msg)
            {
                if (string.IsNullOrWhiteSpace(msg)) return;
                _lines.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] WARN: {msg}");
            }

            public void Error(string msg)
            {
                if (string.IsNullOrWhiteSpace(msg)) return;
                _lines.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {msg}");
            }

            public string TryWrite(string folder, string baseName)
            {
                try
                {
                    if (_lines.Count == 0) return null;
                    string path = Path.Combine(folder ?? "", baseName + "_nest_log.txt");
                    File.WriteAllLines(path, _lines, Encoding.UTF8);
                    return path;
                }
                catch
                {
                    return null;
                }
            }
        }

        private sealed class PartDefinition
        {
            public string BlockName;
            public BlockRecord Block;

            public string MaterialTag;
            public string MaterialType;

            public double MinX, MinY, MaxX, MaxY;
            public double Width, Height;

            public int Quantity;

            public readonly Dictionary<int, RotatedBounds> RotatedCache = new Dictionary<int, RotatedBounds>();
        }

        private struct RotatedBounds
        {
            public double MinX, MinY, MaxX, MaxY;
            public double Width => MaxX - MinX;
            public double Height => MaxY - MinY;
        }

        private sealed class FreeRect
        {
            public double X, Y, Width, Height;
        }

        private sealed class SheetState
        {
            public int Index;
            public double OriginX, OriginY;
            public int PlacedCount;
            public double UsedArea;
            public List<FreeRect> FreeRects = new List<FreeRect>();
        }

        public static int CountTotalInstances(string sourceDwgPath)
        {
            if (string.IsNullOrWhiteSpace(sourceDwgPath) || !File.Exists(sourceDwgPath))
                return 0;

            CadDocument doc;
            using (var reader = new DwgReader(sourceDwgPath))
                doc = reader.Read();

            int candidate, skipped;
            var parts = LoadPartDefinitions(doc, null, out candidate, out skipped).ToList();
            return parts.Sum(p => Math.Max(1, p.Quantity));
        }

        public static NestRunResult NestOneFile(
            string sourceDwgPath,
            LaserCutRunSettings settings,
            LaserCutProgressForm progress,
            ref int placedOverallRef,
            int totalOverall)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            if (string.IsNullOrWhiteSpace(sourceDwgPath) || !File.Exists(sourceDwgPath))
                throw new FileNotFoundException("DWG not found.", sourceDwgPath);

            var log = new NestLog();

            string dir = Path.GetDirectoryName(sourceDwgPath) ?? "";
            string baseName = Path.GetFileNameWithoutExtension(sourceDwgPath) ?? "thickness";

            CadDocument srcDoc;
            using (var reader = new DwgReader(sourceDwgPath))
                srcDoc = reader.Read();

            int candidateBlocks, skippedBlocks;
            var srcParts = LoadPartDefinitions(srcDoc, log, out candidateBlocks, out skippedBlocks).ToList();

            var runResult = new NestRunResult
            {
                SourceDwgPath = sourceDwgPath,
                Settings = settings,
                CandidateBlocks = candidateBlocks,
                SkippedBlocks = skippedBlocks
            };

            if (srcParts.Count == 0)
            {
                log.Warn("No P_* blocks found. Is this a combined thickness DWG from Combine DWG?");
                runResult.LogPath = log.TryWrite(dir, baseName);
                runResult.LogLines = log.Count;
                return runResult;
            }

            List<IGrouping<string, PartDefinition>> groups;

            if (settings.SeparateByMaterial)
            {
                groups = srcParts
                    .GroupBy(p => p.MaterialType ?? MaterialTypeNormalizer.TYPE_OTHER, StringComparer.OrdinalIgnoreCase)
                    .OrderBy(g => MaterialTypeNormalizer.SortKey(g.Key))
                    .ToList();
            }
            else
            {
                groups = new List<IGrouping<string, PartDefinition>>
                {
                    new SingleGrouping<string, PartDefinition>("ALL", srcParts)
                };
            }

            foreach (var g in groups)
            {
                string matType = g.Key;
                SheetSize sheet = settings.GetSheetForMaterialType(matType);

                bool outputPerMaterial = settings.SeparateByMaterial && settings.OutputOneDwgPerMaterial;
                string safeMat = MakeSafeFilePart(matType);

                string outPath =
                    outputPerMaterial
                        ? Path.Combine(dir, $"{baseName}_nested_{safeMat}.dwg")
                        : Path.Combine(dir, $"{baseName}_nested.dwg");

                CadDocument outDoc;
                using (var reader = new DwgReader(sourceDwgPath))
                    outDoc = reader.Read();

                var modelSpace = outDoc.BlockRecords["*Model_Space"];

                object layerSource = EnsureLayer(outDoc, "SOURCE");
                object layerSheet = EnsureLayer(outDoc, "SHEET");
                object layerParts = EnsureLayer(outDoc, "PARTS");
                object layerLabels = EnsureLayer(outDoc, "LABELS");

                // Option 2: filter SOURCE preview to only this material, but KEEP the labels
                if (settings.SeparateByMaterial && !string.Equals(matType, "ALL", StringComparison.OrdinalIgnoreCase))
                {
                    FilterSourcePreviewToMaterial(modelSpace, matType, log);
                }

                foreach (var ent in modelSpace.Entities)
                    SetEntityLayer(ent, layerSource, "SOURCE");

                var srcExt = GetEntityExtents(modelSpace.Entities);
                double baseOriginX = srcExt.HasValue ? srcExt.Value.MinX : 0.0;
                double baseOriginY = srcExt.HasValue ? (srcExt.Value.MaxY + 200.0) : 0.0;

                int dummy1, dummy2;
                var outPartsAll = LoadPartDefinitions(outDoc, log, out dummy1, out dummy2).ToList();

                var outParts = settings.SeparateByMaterial
                    ? outPartsAll.Where(p => string.Equals(p.MaterialType, matType, StringComparison.OrdinalIgnoreCase)).ToList()
                    : outPartsAll;

                if (outParts.Count == 0)
                {
                    log.Warn($"No parts for group '{matType}'. Skipping output.");
                    continue;
                }

                int sheetsUsed;
                int totalParts;

                NestIntoDocument(
                    sourceDwgPath,
                    outDoc,
                    modelSpace,
                    outParts,
                    sheet,
                    baseOriginX,
                    baseOriginY,
                    matType,
                    log,
                    progress,
                    ref placedOverallRef,
                    totalOverall,
                    layerSheet,
                    layerParts,
                    layerLabels,
                    out sheetsUsed,
                    out totalParts);

                using (var writer = new DwgWriter(outPath, outDoc))
                    writer.Write();

                runResult.Outputs.Add(new MaterialNestResult
                {
                    MaterialType = matType,
                    Sheet = sheet,
                    OutputDwgPath = outPath,
                    SheetsUsed = sheetsUsed,
                    TotalParts = totalParts
                });

                if (!outputPerMaterial)
                    break;
            }

            runResult.LogPath = log.TryWrite(dir, baseName);
            runResult.LogLines = log.Count;
            return runResult;
        }

        // ---------------- SOURCE filter (Option 2) ----------------
        // NOTE: We keep only MText labels (Combine writes labels as MText).
        // We DO NOT reference ACadSharp.Entities.Text because your ACadSharp build doesn't have it.

        private static void FilterSourcePreviewToMaterial(BlockRecord modelSpace, string materialType, NestLog log)
        {
            if (modelSpace?.Entities == null || modelSpace.Entities.Count == 0)
                return;

            var ents = modelSpace.Entities.ToList();

            var keepInsert = new HashSet<Insert>();
            var keepBoxes = new List<(double MinX, double MaxX, double MinY)>();

            foreach (var e in ents)
            {
                if (!(e is Insert ins))
                    continue;

                string blockName = TryGetInsertBlockName(ins);
                if (!IsPlateBlock(blockName))
                    continue;

                string tag = ParseMaterialTagFromBlockName(blockName);
                string type = MaterialTypeNormalizer.NormalizeToType((tag ?? "").Replace('_', ' '));

                if (!string.Equals(type, materialType, StringComparison.OrdinalIgnoreCase))
                    continue;

                keepInsert.Add(ins);

                if (TryGetBoundingBox2D(ins, out double minX, out double maxX, out double minY))
                    keepBoxes.Add((minX, maxX, minY));
            }

            // IMPORTANT: if nothing matched, do NOT wipe preview
            if (keepInsert.Count == 0)
            {
                log.Warn($"SOURCE filter: no matching inserts for material '{materialType}'. Preview left unchanged (likely missing __MAT_ tags).");
                return;
            }

            const double padX = 800.0;  // allow long labels
            const double below = 300.0;
            const double above = 120.0;

            bool haveBoxes = keepBoxes.Count > 0;

            bool IsLabelNearAnyKeptInsert(double x, double y)
            {
                if (!haveBoxes)
                {
                    // fallback: keep labels that are below the plates region
                    return y < 200.0;
                }

                foreach (var b in keepBoxes)
                {
                    if (x >= (b.MinX - padX) && x <= (b.MaxX + padX) &&
                        y >= (b.MinY - below) && y <= (b.MinY + above))
                        return true;
                }
                return false;
            }

            var kept = new List<Entity>();

            foreach (var e in ents)
            {
                if (e == null)
                    continue;

                if (e is Insert ins)
                {
                    if (keepInsert.Contains(ins))
                        kept.Add(e);
                    continue;
                }

                if (e is MText)
                {
                    if (TryGetEntityXY(e, out double x, out double y) && IsLabelNearAnyKeptInsert(x, y))
                        kept.Add(e);
                    continue;
                }

                // Remove everything else from SOURCE preview
            }

            modelSpace.Entities.Clear();
            foreach (var e in kept)
                modelSpace.Entities.Add(e);

            log.Info($"SOURCE filter for '{materialType}': kept {kept.Count} entities (inserts={keepInsert.Count}).");
        }

        private static bool IsPlateBlock(string blockName)
        {
            return !string.IsNullOrWhiteSpace(blockName) &&
                   blockName.StartsWith("P_", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryGetBoundingBox2D(Entity ent, out double minX, out double maxX, out double minY)
        {
            minX = maxX = minY = 0.0;

            if (ent == null)
                return false;

            try
            {
                var bb = ent.GetBoundingBox();
                minX = bb.Min.X;
                maxX = bb.Max.X;
                minY = bb.Min.Y;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetEntityXY(Entity ent, out double x, out double y)
        {
            x = y = 0.0;

            try
            {
                if (ent is MText mt)
                {
                    x = mt.InsertPoint.X;
                    y = mt.InsertPoint.Y;
                    return true;
                }
            }
            catch { }

            // Reflection fallback
            try
            {
                var t = ent.GetType();
                var p = t.GetProperty("InsertPoint") ?? t.GetProperty("Position");
                if (p != null)
                {
                    var v = p.GetValue(ent, null);
                    if (v is XYZ xyz)
                    {
                        x = xyz.X;
                        y = xyz.Y;
                        return true;
                    }
                }
            }
            catch { }

            return false;
        }

        private static string TryGetInsertBlockName(Insert ins)
        {
            if (ins == null) return null;

            try
            {
                var pBlock = ins.GetType().GetProperty("Block");
                var blockObj = pBlock?.GetValue(ins);
                if (blockObj != null)
                {
                    var pName = blockObj.GetType().GetProperty("Name");
                    var name = pName?.GetValue(blockObj) as string;
                    if (!string.IsNullOrWhiteSpace(name))
                        return name;
                }
            }
            catch { }

            try
            {
                var p = ins.GetType().GetProperty("BlockName");
                var v = p?.GetValue(ins) as string;
                if (!string.IsNullOrWhiteSpace(v)) return v;
            }
            catch { }

            return null;
        }

        private static string ParseMaterialTagFromBlockName(string blockName)
        {
            if (string.IsNullOrWhiteSpace(blockName)) return "UNKNOWN";

            const string token = "__MAT_";
            int idx = blockName.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return "UNKNOWN";

            int start = idx + token.Length;
            if (start >= blockName.Length) return "UNKNOWN";

            int end = blockName.IndexOf("__", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0) end = blockName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            if (end < 0 || end <= start) end = blockName.Length;

            string tag = blockName.Substring(start, end - start).Trim();
            return string.IsNullOrWhiteSpace(tag) ? "UNKNOWN" : tag;
        }

        // ---------------- nesting core ----------------

        private struct Extents2D
        {
            public double MinX, MinY, MaxX, MaxY;
        }

        private static Extents2D? GetEntityExtents(IEnumerable<Entity> entities)
        {
            double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
            bool any = false;

            foreach (var ent in entities)
            {
                if (ent == null) continue;
                try
                {
                    var bb = ent.GetBoundingBox();
                    var bmin = bb.Min;
                    var bmax = bb.Max;

                    if (bmin.X < minX) minX = bmin.X;
                    if (bmin.Y < minY) minY = bmin.Y;
                    if (bmax.X > maxX) maxX = bmax.X;
                    if (bmax.Y > maxY) maxY = bmax.Y;

                    any = true;
                }
                catch { }
            }

            if (!any) return null;
            return new Extents2D { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };
        }

        private static double? TryParseThicknessFromFilename(string sourceDwgPath)
        {
            try
            {
                string name = Path.GetFileNameWithoutExtension(sourceDwgPath) ?? "";
                if (!name.StartsWith("thickness_", StringComparison.OrdinalIgnoreCase))
                    return null;

                string t = name.Substring("thickness_".Length);
                t = t.Replace('_', '.');

                if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out double mm) && mm > 0 && mm < 500)
                    return mm;
            }
            catch { }
            return null;
        }

        private static void NestIntoDocument(
            string sourceDwgPath,
            CadDocument doc,
            BlockRecord modelSpace,
            List<PartDefinition> parts,
            SheetSize sheet,
            double baseOriginX,
            double baseOriginY,
            string materialType,
            NestLog log,
            LaserCutProgressForm progress,
            ref int placedOverall,
            int totalOverall,
            object layerSheet,
            object layerParts,
            object layerLabels,
            out int sheetsUsed,
            out int totalParts)
        {
            double sheetMargin = 10.0;
            double partGap = 5.0;

            double? thicknessMm = TryParseThicknessFromFilename(sourceDwgPath);
            if (thicknessMm.HasValue && thicknessMm.Value > 0)
            {
                if (thicknessMm.Value > sheetMargin) sheetMargin = thicknessMm.Value;
                if (thicknessMm.Value > partGap) partGap = thicknessMm.Value;
            }

            double placementMargin = sheetMargin + 2 * partGap;

            double usableW = sheet.W - 2 * placementMargin;
            double usableH = sheet.H - 2 * placementMargin;
            if (usableW <= 0 || usableH <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            totalParts = parts.Sum(p => Math.Max(1, p.Quantity));

            foreach (var p in parts)
            {
                bool fits = false;
                foreach (int a in AllowedAnglesDeg)
                {
                    var rb = GetRotatedBounds(p, a);
                    if (rb.Width + partGap <= usableW + 1e-9 && rb.Height + partGap <= usableH + 1e-9)
                    {
                        fits = true;
                        break;
                    }
                }
                if (!fits)
                    throw new InvalidOperationException($"Part '{p.BlockName}' cannot fit into the selected sheet size.");
            }

            var instances = new List<PartDefinition>();
            foreach (var p in parts)
            {
                int q = Math.Max(1, p.Quantity);
                for (int i = 0; i < q; i++)
                    instances.Add(p);
            }

            instances.Sort((a, b) => (b.Width * b.Height).CompareTo(a.Width * a.Height));

            AddGroupLabel(modelSpace, baseOriginX, baseOriginY, sheet, sheetMargin, materialType, layerLabels);

            double sheetGapX = 80.0;

            var sheets = new List<SheetState>();

            SheetState NewSheet()
            {
                int idx = sheets.Count + 1;
                var s = new SheetState
                {
                    Index = idx,
                    OriginX = baseOriginX + (idx - 1) * (sheet.W + sheetGapX),
                    OriginY = baseOriginY
                };

                DrawSheetOutline(s, sheet, modelSpace, layerSheet);

                s.FreeRects.Add(new FreeRect
                {
                    X = placementMargin,
                    Y = placementMargin,
                    Width = sheet.W - 2 * placementMargin,
                    Height = sheet.H - 2 * placementMargin
                });

                sheets.Add(s);
                return s;
            }

            var curSheet = NewSheet();

            foreach (var inst in instances)
            {
                while (true)
                {
                    if (TryPlaceOnSheet(curSheet, inst, partGap, modelSpace, layerParts, ref placedOverall, totalOverall, progress))
                        break;

                    curSheet = NewSheet();
                }
            }

            double usableArea = usableW * usableH;
            foreach (var s in sheets)
            {
                double fillPct = usableArea > 1e-9 ? (s.UsedArea / usableArea) * 100.0 : 0.0;
                AddSheetLabel(modelSpace, s, sheet, sheetMargin, fillPct, layerLabels);
            }

            sheetsUsed = sheets.Count;
        }

        private static bool TryPlaceOnSheet(
            SheetState sheet,
            PartDefinition part,
            double partGap,
            BlockRecord modelSpace,
            object layerParts,
            ref int placedOverall,
            int totalOverall,
            LaserCutProgressForm progress)
        {
            if (sheet.FreeRects.Count == 0)
                return false;

            const double eps = 1e-9;

            int bestRectIndex = -1;
            int bestAngle = 0;
            RotatedBounds bestBounds = default;

            double bestShortSideFit = double.MaxValue;
            double bestLongSideFit = double.MaxValue;
            double bestAreaFit = double.MaxValue;

            for (int i = 0; i < sheet.FreeRects.Count; i++)
            {
                var fr = sheet.FreeRects[i];

                foreach (int ang in AllowedAnglesDeg)
                {
                    var rb = GetRotatedBounds(part, ang);

                    double usedW = rb.Width + partGap;
                    double usedH = rb.Height + partGap;

                    if (usedW > fr.Width + eps || usedH > fr.Height + eps)
                        continue;

                    double leftoverH = fr.Width - usedW;
                    double leftoverV = fr.Height - usedH;

                    double shortFit = Math.Min(leftoverH, leftoverV);
                    double longFit = Math.Max(leftoverH, leftoverV);
                    double areaFit = fr.Width * fr.Height - usedW * usedH;

                    if (shortFit < bestShortSideFit - eps ||
                        (Math.Abs(shortFit - bestShortSideFit) < eps && longFit < bestLongSideFit - eps) ||
                        (Math.Abs(shortFit - bestShortSideFit) < eps && Math.Abs(longFit - bestLongSideFit) < eps && areaFit < bestAreaFit - eps))
                    {
                        bestShortSideFit = shortFit;
                        bestLongSideFit = longFit;
                        bestAreaFit = areaFit;

                        bestRectIndex = i;
                        bestAngle = ang;
                        bestBounds = rb;
                    }
                }
            }

            if (bestRectIndex < 0)
                return false;

            var chosen = sheet.FreeRects[bestRectIndex];

            double usedX = chosen.X;
            double usedY = chosen.Y;

            double partMinLocalX = usedX + partGap * 0.5;
            double partMinLocalY = usedY + partGap * 0.5;

            double worldMinX = sheet.OriginX + partMinLocalX;
            double worldMinY = sheet.OriginY + partMinLocalY;

            double insertX = worldMinX - bestBounds.MinX;
            double insertY = worldMinY - bestBounds.MinY;

            double rotRad = bestAngle * Math.PI / 180.0;

            var insert = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insertX, insertY, 0.0),
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0,
                Rotation = rotRad
            };

            SetEntityLayer(insert, layerParts, "PARTS");
            modelSpace.Entities.Add(insert);

            sheet.PlacedCount++;
            sheet.UsedArea += bestBounds.Width * bestBounds.Height;

            var usedRect = new FreeRect
            {
                X = usedX,
                Y = usedY,
                Width = bestBounds.Width + partGap,
                Height = bestBounds.Height + partGap
            };

            SubtractUsedRect(sheet, usedRect);

            placedOverall++;
            progress?.Step($"Placed {placedOverall}/{totalOverall}");

            return true;
        }

        private static void SubtractUsedRect(SheetState sheet, FreeRect used)
        {
            const double eps = 1e-9;
            const double minSize = 1.0;

            var newFree = new List<FreeRect>();

            double usedRight = used.X + used.Width;
            double usedTop = used.Y + used.Height;

            foreach (var fr in sheet.FreeRects)
            {
                if (!Intersects(fr, used))
                {
                    newFree.Add(fr);
                    continue;
                }

                double frRight = fr.X + fr.Width;
                double frTop = fr.Y + fr.Height;

                if (used.X > fr.X + eps)
                    newFree.Add(new FreeRect { X = fr.X, Y = fr.Y, Width = used.X - fr.X, Height = fr.Height });

                if (usedRight < frRight - eps)
                    newFree.Add(new FreeRect { X = usedRight, Y = fr.Y, Width = frRight - usedRight, Height = fr.Height });

                if (used.Y > fr.Y + eps)
                    newFree.Add(new FreeRect { X = fr.X, Y = fr.Y, Width = fr.Width, Height = used.Y - fr.Y });

                if (usedTop < frTop - eps)
                    newFree.Add(new FreeRect { X = fr.X, Y = usedTop, Width = fr.Width, Height = frTop - usedTop });
            }

            sheet.FreeRects = newFree.Where(r => r.Width > minSize && r.Height > minSize).ToList();

            PruneContained(sheet.FreeRects);
            MergeAdjacent(sheet.FreeRects);
        }

        private static bool Intersects(FreeRect a, FreeRect b)
        {
            return !(a.X + a.Width <= b.X ||
                     b.X + b.Width <= a.X ||
                     a.Y + a.Height <= b.Y ||
                     b.Y + b.Height <= a.Y);
        }

        private static void PruneContained(List<FreeRect> rects)
        {
            const double eps = 1e-9;

            for (int i = rects.Count - 1; i >= 0; i--)
            {
                var a = rects[i];
                for (int j = 0; j < rects.Count; j++)
                {
                    if (i == j) continue;

                    var b = rects[j];

                    bool contained =
                        a.X >= b.X - eps &&
                        a.Y >= b.Y - eps &&
                        a.X + a.Width <= b.X + b.Width + eps &&
                        a.Y + a.Height <= b.Y + b.Height + eps;

                    if (contained)
                    {
                        rects.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        private static void MergeAdjacent(List<FreeRect> rects)
        {
            const double eps = 1e-9;
            bool changed;

            do
            {
                changed = false;

                for (int i = 0; i < rects.Count && !changed; i++)
                {
                    for (int j = i + 1; j < rects.Count && !changed; j++)
                    {
                        var a = rects[i];
                        var b = rects[j];

                        bool sameX = Math.Abs(a.X - b.X) < eps && Math.Abs(a.Width - b.Width) < eps;
                        if (sameX)
                        {
                            if (Math.Abs(a.Y + a.Height - b.Y) < eps)
                            {
                                rects[i] = new FreeRect { X = a.X, Y = a.Y, Width = a.Width, Height = a.Height + b.Height };
                                rects.RemoveAt(j);
                                changed = true;
                                break;
                            }
                            if (Math.Abs(b.Y + b.Height - a.Y) < eps)
                            {
                                rects[i] = new FreeRect { X = b.X, Y = b.Y, Width = b.Width, Height = b.Height + a.Height };
                                rects.RemoveAt(j);
                                changed = true;
                                break;
                            }
                        }

                        bool sameY = Math.Abs(a.Y - b.Y) < eps && Math.Abs(a.Height - b.Height) < eps;
                        if (sameY)
                        {
                            if (Math.Abs(a.X + a.Width - b.X) < eps)
                            {
                                rects[i] = new FreeRect { X = a.X, Y = a.Y, Width = a.Width + b.Width, Height = a.Height };
                                rects.RemoveAt(j);
                                changed = true;
                                break;
                            }
                            if (Math.Abs(b.X + b.Width - a.X) < eps)
                            {
                                rects[i] = new FreeRect { X = b.X, Y = b.Y, Width = b.Width + a.Width, Height = b.Height };
                                rects.RemoveAt(j);
                                changed = true;
                                break;
                            }
                        }
                    }
                }
            } while (changed);
        }

        private static RotatedBounds GetRotatedBounds(PartDefinition part, int angleDeg)
        {
            if (part.RotatedCache.TryGetValue(angleDeg, out var cached))
                return cached;

            double rad = angleDeg * Math.PI / 180.0;
            double c = Math.Cos(rad);
            double s = Math.Sin(rad);

            var pts = new[]
            {
                new XYZ(part.MinX, part.MinY, 0),
                new XYZ(part.MinX, part.MaxY, 0),
                new XYZ(part.MaxX, part.MinY, 0),
                new XYZ(part.MaxX, part.MaxY, 0),
            };

            double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;

            foreach (var p in pts)
            {
                double xr = p.X * c - p.Y * s;
                double yr = p.X * s + p.Y * c;

                if (xr < minX) minX = xr;
                if (yr < minY) minY = yr;
                if (xr > maxX) maxX = xr;
                if (yr > maxY) maxY = yr;
            }

            var rb = new RotatedBounds { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };
            part.RotatedCache[angleDeg] = rb;
            return rb;
        }

        private static void DrawSheetOutline(SheetState sheet, SheetSize sheetSize, BlockRecord modelSpace, object layerSheet)
        {
            var bottom = new Line { StartPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0), EndPoint = new XYZ(sheet.OriginX + sheetSize.W, sheet.OriginY, 0) };
            var right = new Line { StartPoint = new XYZ(sheet.OriginX + sheetSize.W, sheet.OriginY, 0), EndPoint = new XYZ(sheet.OriginX + sheetSize.W, sheet.OriginY + sheetSize.H, 0) };
            var top = new Line { StartPoint = new XYZ(sheet.OriginX + sheetSize.W, sheet.OriginY + sheetSize.H, 0), EndPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetSize.H, 0) };
            var left = new Line { StartPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetSize.H, 0), EndPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0) };

            SetEntityLayer(bottom, layerSheet, "SHEET");
            SetEntityLayer(right, layerSheet, "SHEET");
            SetEntityLayer(top, layerSheet, "SHEET");
            SetEntityLayer(left, layerSheet, "SHEET");

            modelSpace.Entities.Add(bottom);
            modelSpace.Entities.Add(right);
            modelSpace.Entities.Add(top);
            modelSpace.Entities.Add(left);
        }

        private static void AddSheetLabel(BlockRecord modelSpace, SheetState sheet, SheetSize sheetSize, double sheetMargin, double fillPercent, object layerLabels)
        {
            double x = sheet.OriginX + sheetMargin;
            double y = sheet.OriginY + sheetSize.H + sheetMargin;

            var mt = new MText
            {
                Value = $"Sheet {sheet.Index} | Parts: {sheet.PlacedCount} | Fill: {fillPercent:0.0}%",
                InsertPoint = new XYZ(x, y, 0),
                Height = 25.0
            };

            SetEntityLayer(mt, layerLabels, "LABELS");
            modelSpace.Entities.Add(mt);
        }

        private static void AddGroupLabel(BlockRecord modelSpace, double originX, double originY, SheetSize sheetSize, double sheetMargin, string materialType, object layerLabels)
        {
            var mt = new MText
            {
                Value = $"NESTED MATERIAL: {materialType}",
                InsertPoint = new XYZ(originX + sheetMargin, originY + sheetSize.H + sheetMargin + 40.0, 0),
                Height = 35.0
            };

            SetEntityLayer(mt, layerLabels, "LABELS");
            modelSpace.Entities.Add(mt);
        }

        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc, NestLog log, out int candidateBlocks, out int skippedBlocks)
        {
            candidateBlocks = 0;
            skippedBlocks = 0;

            var list = new List<PartDefinition>();

            foreach (var br in doc.BlockRecords)
            {
                if (string.IsNullOrWhiteSpace(br?.Name))
                    continue;

                if (br.Name.StartsWith("*", StringComparison.Ordinal))
                    continue;

                if (!br.Name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                candidateBlocks++;

                int qty = ParseQuantityFromBlockName(br.Name);
                if (qty <= 0) qty = 1;

                string matTag = ParseMaterialTagFromBlockName(br.Name);
                string matType = MaterialTypeNormalizer.NormalizeToType((matTag ?? "").Replace('_', ' '));

                if (br.Entities == null || br.Entities.Count == 0)
                {
                    skippedBlocks++;
                    log?.Warn($"Skipped block '{br.Name}' (empty).");
                    continue;
                }

                double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;

                foreach (var ent in br.Entities)
                {
                    try
                    {
                        var bb = ent.GetBoundingBox();
                        var bmin = bb.Min;
                        var bmax = bb.Max;

                        if (bmin.X < minX) minX = bmin.X;
                        if (bmin.Y < minY) minY = bmin.Y;
                        if (bmax.X > maxX) maxX = bmax.X;
                        if (bmax.Y > maxY) maxY = bmax.Y;
                    }
                    catch { }
                }

                if (minX == double.MaxValue || maxX == double.MinValue || minY == double.MaxValue || maxY == double.MinValue)
                {
                    skippedBlocks++;
                    log?.Warn($"Skipped block '{br.Name}' (no bbox).");
                    continue;
                }

                double w = maxX - minX;
                double h = maxY - minY;
                if (w <= 0 || h <= 0)
                {
                    skippedBlocks++;
                    log?.Warn($"Skipped block '{br.Name}' (invalid size {w:0.###} x {h:0.###}).");
                    continue;
                }

                list.Add(new PartDefinition
                {
                    BlockName = br.Name,
                    Block = br,
                    MaterialTag = matTag,
                    MaterialType = matType,
                    MinX = minX,
                    MinY = minY,
                    MaxX = maxX,
                    MaxY = maxY,
                    Width = w,
                    Height = h,
                    Quantity = qty
                });
            }

            return list;
        }

        private static int ParseQuantityFromBlockName(string blockName)
        {
            if (string.IsNullOrEmpty(blockName)) return 1;

            int idx = blockName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            if (idx < 0 || idx + 2 >= blockName.Length) return 1;

            string s = blockName.Substring(idx + 2);
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int q) && q > 0) return q;
            return 1;
        }

        private static string MakeSafeFilePart(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "OTHER";

            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(s.Length);

            foreach (char c in s.Trim())
            {
                if (invalid.Contains(c)) sb.Append('_');
                else if (char.IsWhiteSpace(c)) sb.Append('_');
                else sb.Append(char.ToUpperInvariant(c));
            }

            string r = sb.ToString().Trim('_');
            return string.IsNullOrWhiteSpace(r) ? "OTHER" : r;
        }

        // -------- Layers (best-effort, reflection-safe) --------

        private static object EnsureLayer(CadDocument doc, string layerName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(layerName))
                return null;

            try
            {
                var layersProp = doc.GetType().GetProperty("Layers");
                var layersObj = layersProp?.GetValue(doc);
                if (layersObj == null)
                    return null;

                var indexer = layersObj.GetType().GetProperty("Item", new[] { typeof(string) });
                if (indexer != null)
                {
                    try
                    {
                        var existing = indexer.GetValue(layersObj, new object[] { layerName });
                        if (existing != null)
                            return existing;
                    }
                    catch { }
                }

                var newLayer = new Layer(layerName);

                var addAny = layersObj.GetType().GetMethods()
                    .FirstOrDefault(m => m.Name == "Add" && m.GetParameters().Length == 1);

                addAny?.Invoke(layersObj, new object[] { newLayer });
                return newLayer;
            }
            catch
            {
                return null;
            }
        }

        private static void SetEntityLayer(Entity ent, object layerObj, string layerName)
        {
            if (ent == null) return;

            try
            {
                var pLayer = ent.GetType().GetProperty("Layer");
                if (pLayer != null && pLayer.CanWrite)
                {
                    if (pLayer.PropertyType == typeof(string))
                    {
                        pLayer.SetValue(ent, layerName);
                        return;
                    }

                    if (layerObj != null && pLayer.PropertyType.IsInstanceOfType(layerObj))
                    {
                        pLayer.SetValue(ent, layerObj);
                        return;
                    }
                }

                var pLayerName = ent.GetType().GetProperty("LayerName");
                if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
                {
                    pLayerName.SetValue(ent, layerName);
                }
            }
            catch { }
        }

        private sealed class SingleGrouping<TKey, TElement> : IGrouping<TKey, TElement>
        {
            private readonly IEnumerable<TElement> _elements;
            public TKey Key { get; }

            public SingleGrouping(TKey key, IEnumerable<TElement> elements)
            {
                Key = key;
                _elements = elements ?? Enumerable.Empty<TElement>();
            }

            public IEnumerator<TElement> GetEnumerator() => _elements.GetEnumerator();
            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }

    // ---------------- Material normalization ----------------
    internal static class MaterialTypeNormalizer
    {
        public const string TYPE_STEEL = "STEEL";
        public const string TYPE_ALUMINUM = "ALUMINUM";
        public const string TYPE_STAINLESS = "STAINLESS";
        public const string TYPE_OTHER = "OTHER";

        public static string NormalizeToType(string materialRawOrTag)
        {
            if (string.IsNullOrWhiteSpace(materialRawOrTag))
                return TYPE_OTHER;

            string s = materialRawOrTag.Trim().Replace('_', ' ').ToUpperInvariant();
            while (s.Contains("  ")) s = s.Replace("  ", " ");

            // Stainless first
            if (ContainsAny(s,
                "STAINLESS", "INOX", "SUS",
                "AISI 304", "AISI304", "SS304", "SS 304", "304",
                "AISI 316", "AISI316", "SS316", "SS 316", "316",
                "1.4301", "1.4401", "1.4404", "1.4571",
                "A2", "A4"))
                return TYPE_STAINLESS;

            if (ContainsAny(s,
                "ALUMIN", "ALUMINUM", "ALU", "AL ",
                "6061", "6082", "7075", "5052", "5083"))
                return TYPE_ALUMINUM;

            if (ContainsAny(s,
                "STEEL", "S235", "S355", "ST37", "ST52", "A36", "C45", "CK45"))
                return TYPE_STEEL;

            return TYPE_OTHER;
        }

        public static int SortKey(string type)
        {
            type = (type ?? "").Trim().ToUpperInvariant();
            if (type == TYPE_STEEL) return 1;
            if (type == TYPE_ALUMINUM) return 2;
            if (type == TYPE_STAINLESS) return 3;
            return 4;
        }

        private static bool ContainsAny(string s, params string[] keys)
        {
            foreach (var k in keys)
            {
                if (string.IsNullOrWhiteSpace(k)) continue;
                if (s.Contains(k)) return true;
            }
            return false;
        }
    }

    // ---------------- UI (Options) ----------------
    internal sealed class LaserCutOptionsForm : Form
    {
        private sealed class SheetPreset
        {
            public string Name;
            public SheetSize Size;
            public bool IsCustom;
            public override string ToString() => Name;
        }

        private static class UiSettings
        {
            private const string RegPath = @"Software\SW2026RibbonAddin\LaserCut";

            public static int LoadInt(string name, int def)
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(RegPath, false))
                    {
                        object v = k?.GetValue(name);
                        if (v == null) return def;
                        if (v is int i) return i;
                        if (int.TryParse(v.ToString(), out i)) return i;
                    }
                }
                catch { }
                return def;
            }

            public static double LoadDouble(string name, double def)
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(RegPath, false))
                    {
                        object v = k?.GetValue(name);
                        if (v == null) return def;
                        if (double.TryParse(v.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) return d;
                    }
                }
                catch { }
                return def;
            }

            public static bool LoadBool(string name, bool def)
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(RegPath, false))
                    {
                        object v = k?.GetValue(name);
                        if (v == null) return def;
                        if (v is int i) return i != 0;
                        if (bool.TryParse(v.ToString(), out bool b)) return b;
                    }
                }
                catch { }
                return def;
            }

            public static void Save(Dictionary<string, object> values)
            {
                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(RegPath))
                    {
                        foreach (var kvp in values)
                        {
                            if (kvp.Value is int i) k.SetValue(kvp.Key, i, RegistryValueKind.DWord);
                            else if (kvp.Value is bool b) k.SetValue(kvp.Key, b ? 1 : 0, RegistryValueKind.DWord);
                            else if (kvp.Value is double d) k.SetValue(kvp.Key, d.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String);
                            else if (kvp.Value is string s) k.SetValue(kvp.Key, s, RegistryValueKind.String);
                        }
                    }
                }
                catch { }
            }
        }

        private readonly List<SheetPreset> _presets;
        private readonly List<SheetPreset> _materialPresets;

        private readonly ComboBox _cmbGlobalPreset;
        private readonly TextBox _txtW;
        private readonly TextBox _txtH;

        private readonly CheckBox _chkSeparate;
        private readonly CheckBox _chkOneDwgPerMaterial;
        private readonly CheckBox _chkUsePerMaterialSheets;

        private readonly ComboBox _cmbSteel;
        private readonly ComboBox _cmbAlu;
        private readonly ComboBox _cmbStainless;
        private readonly ComboBox _cmbOther;

        public LaserCutRunSettings Settings { get; private set; }

        public LaserCutOptionsForm()
        {
            Text = "Batch nest options (applies to ALL thickness files)";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;

            ClientSize = new WinSize(580, 390);

            _presets = new List<SheetPreset>
            {
                new SheetPreset { Name = "1500 x 3000 mm", Size = new SheetSize(3000, 1500), IsCustom = false },
                new SheetPreset { Name = "1250 x 2500 mm", Size = new SheetSize(2500, 1250), IsCustom = false },
                new SheetPreset { Name = "1000 x 2000 mm", Size = new SheetSize(2000, 1000), IsCustom = false },
                new SheetPreset { Name = "2000 x 4000 mm", Size = new SheetSize(4000, 2000), IsCustom = false },
                new SheetPreset { Name = "Custom...", Size = new SheetSize(3000, 1500), IsCustom = true },
            };

            _materialPresets = _presets.Where(p => !p.IsCustom).ToList();

            Controls.Add(new Label { AutoSize = true, Text = "Global sheet preset:", Location = new WinPoint(12, 18) });

            _cmbGlobalPreset = new ComboBox
            {
                Location = new WinPoint(170, 14),
                Width = 260,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbGlobalPreset.Items.AddRange(_presets.Cast<object>().ToArray());
            _cmbGlobalPreset.SelectedIndexChanged += (s, e) => ApplyGlobalPresetToFields();
            Controls.Add(_cmbGlobalPreset);

            Controls.Add(new Label { AutoSize = true, Text = "Width (mm):", Location = new WinPoint(12, 54) });
            _txtW = new TextBox { Location = new WinPoint(170, 50), Width = 120 };
            Controls.Add(_txtW);

            Controls.Add(new Label { AutoSize = true, Text = "Height (mm):", Location = new WinPoint(12, 88) });
            _txtH = new TextBox { Location = new WinPoint(170, 84), Width = 120 };
            Controls.Add(_txtH);

            _chkSeparate = new CheckBox
            {
                AutoSize = true,
                Text = "Separate nests by material type (STEEL / ALUMINUM / STAINLESS / OTHER)",
                Location = new WinPoint(12, 124)
            };
            _chkSeparate.CheckedChanged += (s, e) => UpdateMaterialUiEnabled();
            Controls.Add(_chkSeparate);

            _chkOneDwgPerMaterial = new CheckBox
            {
                AutoSize = true,
                Text = "Output one DWG per material type",
                Location = new WinPoint(32, 152)
            };
            Controls.Add(_chkOneDwgPerMaterial);

            _chkUsePerMaterialSheets = new CheckBox
            {
                AutoSize = true,
                Text = "Use per-material sheet presets",
                Location = new WinPoint(32, 178)
            };
            _chkUsePerMaterialSheets.CheckedChanged += (s, e) => UpdateMaterialUiEnabled();
            Controls.Add(_chkUsePerMaterialSheets);

            var grp = new GroupBox
            {
                Text = "Per-material sheet preset (when enabled)",
                Location = new WinPoint(12, 210),
                Size = new WinSize(556, 90)
            };
            Controls.Add(grp);

            grp.Controls.Add(new Label { AutoSize = true, Text = "Steel:", Location = new WinPoint(12, 28) });
            _cmbSteel = MakeMaterialCombo();
            _cmbSteel.Location = new WinPoint(60, 24);
            grp.Controls.Add(_cmbSteel);

            grp.Controls.Add(new Label { AutoSize = true, Text = "Alu:", Location = new WinPoint(220, 28) });
            _cmbAlu = MakeMaterialCombo();
            _cmbAlu.Location = new WinPoint(250, 24);
            grp.Controls.Add(_cmbAlu);

            grp.Controls.Add(new Label { AutoSize = true, Text = "SS:", Location = new WinPoint(410, 28) });
            _cmbStainless = MakeMaterialCombo();
            _cmbStainless.Location = new WinPoint(440, 24);
            grp.Controls.Add(_cmbStainless);

            grp.Controls.Add(new Label { AutoSize = true, Text = "Other:", Location = new WinPoint(12, 58) });
            _cmbOther = MakeMaterialCombo();
            _cmbOther.Location = new WinPoint(60, 54);
            grp.Controls.Add(_cmbOther);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 54 };
            Controls.Add(bottom);

            var note = new Label
            {
                AutoSize = false,
                Location = new WinPoint(12, 8),
                Size = new WinSize(360, 40),
                Text = "Note: rotations are always 0/90/180/270.\r\nGap+margin are auto (>= thickness)."
            };
            note.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            bottom.Controls.Add(note);

            var btnOk = new Button { Text = "OK", Width = 85, Height = 28 };
            btnOk.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnOk.Location = new WinPoint(bottom.Width - 190, 12);
            btnOk.Click += Ok_Click;
            bottom.Controls.Add(btnOk);

            var btnCancel = new Button { Text = "Cancel", Width = 85, Height = 28, DialogResult = DialogResult.Cancel };
            btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnCancel.Location = new WinPoint(bottom.Width - 95, 12);
            bottom.Controls.Add(btnCancel);

            bottom.Resize += (s, e) =>
            {
                btnOk.Location = new WinPoint(bottom.Width - 190, 12);
                btnCancel.Location = new WinPoint(bottom.Width - 95, 12);
                note.Size = new WinSize(Math.Max(100, bottom.Width - 210), 40);
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            int globalIdx = UiSettings.LoadInt("GlobalPresetIndex", 0);
            if (globalIdx < 0 || globalIdx >= _presets.Count) globalIdx = 0;

            double customW = UiSettings.LoadDouble("CustomW", 3000);
            double customH = UiSettings.LoadDouble("CustomH", 1500);

            bool separate = UiSettings.LoadBool("SeparateByMaterial", true);
            bool onePerMat = UiSettings.LoadBool("OneDwgPerMaterial", true);
            bool perMatSheets = UiSettings.LoadBool("UsePerMaterialSheets", true);

            _cmbGlobalPreset.SelectedIndex = globalIdx;

            _txtW.Text = customW.ToString("0.###", CultureInfo.InvariantCulture);
            _txtH.Text = customH.ToString("0.###", CultureInfo.InvariantCulture);

            _chkSeparate.Checked = separate;
            _chkOneDwgPerMaterial.Checked = onePerMat;
            _chkUsePerMaterialSheets.Checked = perMatSheets;

            int steelIdx = Clamp(UiSettings.LoadInt("SteelPresetIndex", 0), 0, _materialPresets.Count - 1);
            int aluIdx = Clamp(UiSettings.LoadInt("AluPresetIndex", 1), 0, _materialPresets.Count - 1);
            int ssIdx = Clamp(UiSettings.LoadInt("StainlessPresetIndex", 0), 0, _materialPresets.Count - 1);
            int otherIdx = Clamp(UiSettings.LoadInt("OtherPresetIndex", 0), 0, _materialPresets.Count - 1);

            _cmbSteel.SelectedIndex = steelIdx;
            _cmbAlu.SelectedIndex = aluIdx;
            _cmbStainless.SelectedIndex = ssIdx;
            _cmbOther.SelectedIndex = otherIdx;

            ApplyGlobalPresetToFields();
            UpdateMaterialUiEnabled();
        }

        private ComboBox MakeMaterialCombo()
        {
            var cmb = new ComboBox
            {
                Width = 130,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmb.Items.AddRange(_materialPresets.Cast<object>().ToArray());
            return cmb;
        }

        private static int Clamp(int v, int lo, int hi)
        {
            if (v < lo) return lo;
            if (v > hi) return hi;
            return v;
        }

        private void ApplyGlobalPresetToFields()
        {
            if (!(_cmbGlobalPreset.SelectedItem is SheetPreset p))
                return;

            if (!p.IsCustom)
            {
                _txtW.Text = p.Size.W.ToString("0.###", CultureInfo.InvariantCulture);
                _txtH.Text = p.Size.H.ToString("0.###", CultureInfo.InvariantCulture);
                _txtW.Enabled = false;
                _txtH.Enabled = false;
            }
            else
            {
                _txtW.Enabled = true;
                _txtH.Enabled = true;
            }
        }

        private void UpdateMaterialUiEnabled()
        {
            bool sep = _chkSeparate.Checked;

            _chkOneDwgPerMaterial.Enabled = sep;
            _chkUsePerMaterialSheets.Enabled = sep;

            bool enablePresets = sep && _chkUsePerMaterialSheets.Checked;

            _cmbSteel.Enabled = enablePresets;
            _cmbAlu.Enabled = enablePresets;
            _cmbStainless.Enabled = enablePresets;
            _cmbOther.Enabled = enablePresets;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (!(_cmbGlobalPreset.SelectedItem is SheetPreset globalPreset))
                return;

            SheetSize globalSize;

            if (!globalPreset.IsCustom)
            {
                globalSize = globalPreset.Size;
            }
            else
            {
                if (!double.TryParse(_txtW.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double w) || w <= 0)
                {
                    MessageBox.Show(this, "Enter a valid positive width (mm).", "Laser Cut", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _txtW.Focus(); _txtW.SelectAll();
                    return;
                }
                if (!double.TryParse(_txtH.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double h) || h <= 0)
                {
                    MessageBox.Show(this, "Enter a valid positive height (mm).", "Laser Cut", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _txtH.Focus(); _txtH.SelectAll();
                    return;
                }
                globalSize = new SheetSize(w, h);
            }

            bool sep = _chkSeparate.Checked;
            bool onePerMat = sep && _chkOneDwgPerMaterial.Checked;
            bool perMatSheets = sep && _chkUsePerMaterialSheets.Checked;

            SheetSize steel = _materialPresets[_cmbSteel.SelectedIndex].Size;
            SheetSize alu = _materialPresets[_cmbAlu.SelectedIndex].Size;
            SheetSize ss = _materialPresets[_cmbStainless.SelectedIndex].Size;
            SheetSize other = _materialPresets[_cmbOther.SelectedIndex].Size;

            Settings = new LaserCutRunSettings
            {
                SeparateByMaterial = sep,
                OutputOneDwgPerMaterial = onePerMat,
                UsePerMaterialSheetPresets = perMatSheets,

                DefaultSheet = globalSize,

                SteelSheet = steel,
                AluminumSheet = alu,
                StainlessSheet = ss,
                OtherSheet = other
            };

            var values = new Dictionary<string, object>
            {
                ["GlobalPresetIndex"] = _cmbGlobalPreset.SelectedIndex,
                ["CustomW"] = globalSize.W,
                ["CustomH"] = globalSize.H,

                ["SeparateByMaterial"] = sep,
                ["OneDwgPerMaterial"] = onePerMat,
                ["UsePerMaterialSheets"] = perMatSheets,

                ["SteelPresetIndex"] = _cmbSteel.SelectedIndex,
                ["AluPresetIndex"] = _cmbAlu.SelectedIndex,
                ["StainlessPresetIndex"] = _cmbStainless.SelectedIndex,
                ["OtherPresetIndex"] = _cmbOther.SelectedIndex
            };

            UiSettings.Save(values);

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ---------------- UI (Progress) ----------------
    internal sealed class LaserCutProgressForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _label;
        private readonly int _maximum;
        private int _current;

        public LaserCutProgressForm(int maximum)
        {
            if (maximum <= 0) maximum = 1;
            _maximum = maximum;

            Text = "Laser nesting...";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new WinSize(520, 95);

            _label = new Label
            {
                AutoSize = false,
                Text = "Preparing...",
                TextAlign = WinContentAlignment.MiddleLeft,
                Location = new WinPoint(12, 9),
                Size = new WinSize(496, 22)
            };
            Controls.Add(_label);

            _progressBar = new ProgressBar
            {
                Location = new WinPoint(12, 40),
                Size = new WinSize(496, 20),
                Minimum = 0,
                Maximum = _maximum,
                Value = 0
            };
            Controls.Add(_progressBar);
        }

        public void SetStatus(string statusText)
        {
            if (!IsHandleCreated) return;
            if (!string.IsNullOrEmpty(statusText))
                _label.Text = statusText;

            _label.Refresh();
            WinFormsApp.DoEvents();
        }

        public void Step(string statusText)
        {
            if (!IsHandleCreated) return;

            if (!string.IsNullOrEmpty(statusText))
                _label.Text = statusText;

            if (_current < _maximum)
            {
                _current++;
                _progressBar.Value = _current;
            }

            _progressBar.Refresh();
            _label.Refresh();
            WinFormsApp.DoEvents();
        }
    }
}
