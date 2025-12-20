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

        public string DisplayName => "Laser\nCut";
        public string Tooltip => "Nest a combined DWG into laser sheets (improved packing + layers + rotation rules).";
        public string Hint => "Laser cut nesting";

        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string dwgPath = SelectCombinedDwg();
            if (string.IsNullOrEmpty(dwgPath))
                return;

            double sheetWidth;
            double sheetHeight;
            RotationMode rotationMode;
            int anyAngleStepDeg;
            bool writeReportCsv;

            // Ask for sheet size + rotation mode + report option
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                sheetWidth = dlg.SheetWidthMm;
                sheetHeight = dlg.SheetHeightMm;
                rotationMode = dlg.RotationMode;
                anyAngleStepDeg = dlg.AnyAngleStepDeg;
                writeReportCsv = dlg.WriteReportCsv;
            }

            try
            {
                DwgLaserNester.Nest(dwgPath, sheetWidth, sheetHeight, rotationMode, anyAngleStepDeg, writeReportCsv);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Laser cut nesting failed:\r\n\r\n" + ex.Message,
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // Independent of active SW document
            return AddinContext.Enable;
        }

        private static string SelectCombinedDwg()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select combined thickness DWG";
                dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.Multiselect = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.FileName;
            }
        }
    }

    internal enum RotationMode
    {
        Deg0Only = 0,
        Deg0_90 = 1,
        Deg0_90_180_270 = 2,
        AnyAngleStep = 3
    }

    internal static class DwgLaserNester
    {
        private sealed class PartDefinition
        {
            public BlockRecord Block;
            public string BlockName;

            public double MinX;
            public double MinY;
            public double MaxX;
            public double MaxY;

            public double Width;
            public double Height;

            public int Quantity;

            // Cache rotated bounds by integer degrees (for speed)
            public readonly Dictionary<int, RotatedBounds> RotatedCache = new Dictionary<int, RotatedBounds>();
        }

        private struct RotatedBounds
        {
            public double MinX;
            public double MinY;
            public double MaxX;
            public double MaxY;

            public double Width => MaxX - MinX;
            public double Height => MaxY - MinY;
        }

        private sealed class SheetState
        {
            public int Index;
            public double OriginX;
            public double OriginY;

            public double UsedArea;      // sum of bounding areas (rough)
            public int PlacedCount;

            public List<FreeRect> FreeRects = new List<FreeRect>();
        }

        private sealed class FreeRect
        {
            public double X;
            public double Y;
            public double Width;
            public double Height;
        }

        private sealed class PlacementRecord
        {
            public int SheetIndex;
            public string BlockName;
            public int AngleDeg;

            public double LocalMinX;
            public double LocalMinY;
            public double BoundW;
            public double BoundH;

            public double InsertWorldX;
            public double InsertWorldY;
        }

        /// <summary>
        /// Try to determine plate thickness (in mm) from the DWG file name
        /// produced by CombineDwg, e.g. "thickness_2_5.dwg".
        /// </summary>
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

            // Combined DWGs are produced with decimal separators replaced by '_'
            token = token.Replace('_', '.');

            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                value > 0.0 && value < 1000.0)
            {
                return value;
            }

            return null;
        }

        /// <summary>
        /// Try to determine plate thickness (in mm) by parsing MText labels
        /// like "Plate: X mm" that CombineDwg writes under each plate.
        /// </summary>
        private static double? TryGetPlateThicknessFromMText(CadDocument doc)
        {
            if (doc == null)
                return null;

            try
            {
                foreach (var ent in doc.Entities)
                {
                    if (ent is MText mtext)
                    {
                        string text = mtext.Value;
                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        int idx = text.IndexOf("Plate:", StringComparison.OrdinalIgnoreCase);
                        if (idx < 0)
                            continue;

                        string after = text.Substring(idx + "Plate:".Length).Trim();

                        int mmIdx = after.IndexOf("mm", StringComparison.OrdinalIgnoreCase);
                        if (mmIdx >= 0)
                            after = after.Substring(0, mmIdx).Trim();

                        if (string.IsNullOrWhiteSpace(after))
                            continue;

                        after = after.Replace(',', '.');

                        if (double.TryParse(after, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                            value > 0.0 && value < 1000.0)
                        {
                            return value;
                        }
                    }
                }
            }
            catch
            {
                // best-effort only
            }

            return null;
        }

        private static double? TryGetPlateThicknessMm(CadDocument doc, string sourceDwgPath)
        {
            var fromFileName = TryGetPlateThicknessFromFileName(sourceDwgPath);
            if (fromFileName.HasValue)
                return fromFileName;

            return TryGetPlateThicknessFromMText(doc);
        }

        private static List<int> BuildAngleList(RotationMode mode, int stepDeg)
        {
            var angles = new List<int>();

            switch (mode)
            {
                case RotationMode.Deg0Only:
                    angles.Add(0);
                    break;

                case RotationMode.Deg0_90:
                    angles.Add(0);
                    angles.Add(90);
                    break;

                case RotationMode.Deg0_90_180_270:
                    angles.Add(0);
                    angles.Add(90);
                    angles.Add(180);
                    angles.Add(270);
                    break;

                case RotationMode.AnyAngleStep:
                default:
                    if (stepDeg < 1) stepDeg = 10;
                    if (stepDeg > 90) stepDeg = 90;

                    // True “any angle” is infinite; this is practical sampling.
                    // 0..180 is enough for cutting (180..360 is same part orientation flipped).
                    for (int a = 0; a <= 180; a += stepDeg)
                        angles.Add(a);

                    if (!angles.Contains(0)) angles.Insert(0, 0);
                    if (!angles.Contains(90)) angles.Add(90);
                    if (!angles.Contains(180)) angles.Add(180);
                    break;
            }

            // Remove duplicates + keep sorted
            angles = angles.Distinct().OrderBy(x => x).ToList();
            return angles;
        }

        /// <summary>
        /// Compute rotated bounds of the part's bounding rectangle for a given angle (deg).
        /// Rotation is around the insert point / origin (0,0).
        /// We rotate the 4 bbox corners and take min/max.
        /// </summary>
        private static RotatedBounds GetRotatedBounds(PartDefinition part, int angleDeg)
        {
            if (part.RotatedCache.TryGetValue(angleDeg, out RotatedBounds cached))
                return cached;

            double rad = angleDeg * Math.PI / 180.0;
            double c = Math.Cos(rad);
            double s = Math.Sin(rad);

            // bbox corners
            var pts = new[]
            {
                new XYZ(part.MinX, part.MinY, 0),
                new XYZ(part.MinX, part.MaxY, 0),
                new XYZ(part.MaxX, part.MinY, 0),
                new XYZ(part.MaxX, part.MaxY, 0),
            };

            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;

            foreach (var p in pts)
            {
                double x = p.X;
                double y = p.Y;

                double xr = x * c - y * s;
                double yr = x * s + y * c;

                if (xr < minX) minX = xr;
                if (yr < minY) minY = yr;
                if (xr > maxX) maxX = xr;
                if (yr > maxY) maxY = yr;
            }

            var rb = new RotatedBounds
            {
                MinX = minX,
                MinY = minY,
                MaxX = maxX,
                MaxY = maxY
            };

            part.RotatedCache[angleDeg] = rb;
            return rb;
        }

        /// <summary>
        /// Layer support: create layers if the library supports it.
        /// We use reflection to avoid compile breaks if the API differs.
        /// </summary>
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

                // Try string indexer: layers["NAME"]
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

                // Create and add
                var newLayer = new Layer(layerName);

                // Find Add(Layer)
                var add = layersObj.GetType().GetMethods()
                    .FirstOrDefault(m =>
                        m.Name == "Add" &&
                        m.GetParameters().Length == 1 &&
                        m.GetParameters()[0].ParameterType.IsAssignableFrom(typeof(Layer)));

                if (add != null)
                {
                    add.Invoke(layersObj, new object[] { newLayer });
                    return newLayer;
                }

                // Some versions might have Add(object) or Add(TableEntry)
                var addAny = layersObj.GetType().GetMethods()
                    .FirstOrDefault(m => m.Name == "Add" && m.GetParameters().Length == 1);

                if (addAny != null)
                {
                    addAny.Invoke(layersObj, new object[] { newLayer });
                    return newLayer;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static void SetEntityLayer(Entity ent, object layerObj, string layerName)
        {
            if (ent == null)
                return;

            try
            {
                // Common patterns: Entity.Layer (Layer), Entity.Layer (string), or Entity.LayerName (string)
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
                    return;
                }
            }
            catch
            {
                // ignore
            }
        }

        /// <summary>
        /// Reads a combined DWG (from CombineDwg) and nests all plate blocks P_*_Q{qty}
        /// onto as many sheets as required.
        /// Output DWG contains:
        /// - original layout on SOURCE layer
        /// - sheets on SHEET layer
        /// - nested parts on PARTS layer
        /// - labels on LABELS layer
        /// </summary>
        public static void Nest(
            string sourceDwgPath,
            double sheetWidth,
            double sheetHeight,
            RotationMode rotationMode,
            int anyAngleStepDeg,
            bool writeReportCsv)
        {
            if (sheetWidth <= 0 || sheetHeight <= 0)
                throw new ArgumentException("Sheet width and height must be positive.");

            if (!File.Exists(sourceDwgPath))
                throw new FileNotFoundException("DWG file not found.", sourceDwgPath);

            CadDocument doc;
            using (var reader = new DwgReader(sourceDwgPath))
            {
                doc = reader.Read();
            }

            var modelSpace = doc.BlockRecords["*Model_Space"];

            // --- layers (safe; if ACadSharp doesn’t support layers in your build, it just won’t crash) ---
            object layerSource = EnsureLayer(doc, "SOURCE");
            object layerSheet = EnsureLayer(doc, "SHEET");
            object layerParts = EnsureLayer(doc, "PARTS");
            object layerLabels = EnsureLayer(doc, "LABELS");

            // Move the original combined layout to SOURCE layer
            var originalEntities = modelSpace.Entities.ToList();
            foreach (var ent in originalEntities)
                SetEntityLayer(ent, layerSource, "SOURCE");

            var parts = LoadPartDefinitions(doc).ToList();
            if (parts.Count == 0)
            {
                MessageBox.Show(
                    "No plate blocks were found in the selected DWG.\r\n" +
                    "Make sure it is one of the combined thickness DWGs.",
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            int totalInstances = parts.Sum(p => p.Quantity);
            if (totalInstances <= 0)
            {
                MessageBox.Show(
                    "All parts in the selected DWG have zero quantity.",
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Sheet framing and spacing
            double sheetMargin = 10.0;     // visible border (mm)
            double defaultPartGap = 5.0;   // nominal gap between parts (mm)
            double sheetGap = 50.0;        // distance between sheets (mm)

            // Determine plate thickness and ensure gap is never smaller than plate thickness
            double partGap = defaultPartGap;
            double? plateThicknessMm = TryGetPlateThicknessMm(doc, sourceDwgPath);
            if (plateThicknessMm.HasValue && plateThicknessMm.Value > 0)
            {
                if (plateThicknessMm.Value > partGap) partGap = plateThicknessMm.Value;
                if (plateThicknessMm.Value > sheetMargin) sheetMargin = plateThicknessMm.Value;
            }

            // Clearance margin so nothing gets too close to sheet border
            double placementMargin = sheetMargin + 2 * partGap;

            double usableWidth = sheetWidth - 2 * placementMargin;
            double usableHeight = sheetHeight - 2 * placementMargin;

            if (usableWidth <= 0 || usableHeight <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gaps.");

            // Rotation angle list
            var anglesDeg = BuildAngleList(rotationMode, anyAngleStepDeg);

            // Validate: every part must fit on usable area in at least one allowed rotation
            foreach (var p in parts)
            {
                bool fits = false;

                foreach (int a in anglesDeg)
                {
                    var rb = GetRotatedBounds(p, a);
                    double usedW = rb.Width + partGap;
                    double usedH = rb.Height + partGap;

                    if (usedW <= usableWidth + 1e-9 && usedH <= usableHeight + 1e-9)
                    {
                        fits = true;
                        break;
                    }
                }

                if (!fits)
                {
                    throw new InvalidOperationException(
                        $"Part '{p.BlockName}' cannot fit inside sheet {sheetWidth:0.##} x {sheetHeight:0.##} mm\r\n" +
                        $"with margin {placementMargin:0.##} and gap {partGap:0.##} using the selected rotation rules.");
                }
            }

            // Build instances list (expand quantities)
            var instances = new List<PartDefinition>(totalInstances);
            foreach (var def in parts)
            {
                for (int i = 0; i < def.Quantity; i++)
                    instances.Add(def);
            }

            // Place largest parts first (by bbox area)
            instances.Sort((a, b) =>
            {
                double areaA = a.Width * a.Height;
                double areaB = b.Width * b.Height;
                return areaB.CompareTo(areaA);
            });

            // Compute extents of original combined layout to place sheets above it
            GetModelSpaceExtents(doc, out double origMinX, out double origMinY, out double origMaxX, out double origMaxY);

            double baseSheetOriginY = origMaxY + 200.0;
            double baseSheetOriginX = origMinX;

            var progress = new LaserCutProgressForm(totalInstances)
            {
                Text = "Laser cut nesting (improved)"
            };

            List<SheetState> sheets;
            List<PlacementRecord> placements;

            try
            {
                progress.Show();
                Application.DoEvents();

                NestFreeRectangles(
                    instances,
                    modelSpace,
                    sheetWidth,
                    sheetHeight,
                    placementMargin,
                    sheetGap,
                    partGap,
                    baseSheetOriginX,
                    baseSheetOriginY,
                    progress,
                    totalInstances,
                    anglesDeg,
                    layerSheet,
                    layerParts,
                    out sheets,
                    out placements);
            }
            finally
            {
                progress.Close();
            }

            // Add sheet labels (after nesting so we know counts + fill)
            double usableArea = usableWidth * usableHeight;
            foreach (var s in sheets)
            {
                double fillPct = usableArea > 1e-9 ? (s.UsedArea / usableArea) * 100.0 : 0.0;
                AddSheetLabel(modelSpace, s, sheetWidth, sheetHeight, sheetMargin, fillPct, layerLabels);
            }

            // Write nested result to a new DWG next to the source
            string dir = Path.GetDirectoryName(sourceDwgPath);
            string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath);
            string outPath = Path.Combine(dir ?? string.Empty, nameNoExt + "_nested_optimized.dwg");

            using (var writer = new DwgWriter(outPath, doc))
            {
                writer.Write();
            }

            // Optional report CSV
            string reportPath = null;
            if (writeReportCsv)
            {
                try
                {
                    reportPath = Path.Combine(dir ?? string.Empty, nameNoExt + "_nest_report.csv");
                    WriteReportCsv(reportPath, sheetWidth, sheetHeight, placementMargin, partGap, sheets, placements, usableWidth, usableHeight);
                }
                catch
                {
                    reportPath = null;
                }
            }

            MessageBox.Show(
                "Laser cut nesting finished.\r\n\r\n" +
                "Algorithm: Improved (MaxRects-style + prune + merge)\r\n" +
                "Sheets used: " + sheets.Count + "\r\n" +
                "Total parts: " + totalInstances + "\r\n" +
                "Output DWG: " + outPath +
                (string.IsNullOrEmpty(reportPath) ? "" : ("\r\nReport CSV: " + reportPath)),
                "Laser Cut",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        #region Improved packing (MaxRects-style split + prune + merge)

        private static void NestFreeRectangles(
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
            List<int> anglesDeg,
            object layerSheet,
            object layerParts,
            out List<SheetState> sheets,
            out List<PlacementRecord> placements)
        {
            // ✅ Use locals (allowed inside local functions), then assign to out parameters at the end.
            var sheetsLocal = new List<SheetState>();
            var placementsLocal = new List<PlacementRecord>(instances.Count);

            SheetState NewSheet()
            {
                var sheet = new SheetState
                {
                    Index = sheetsLocal.Count + 1,
                    OriginX = startOriginX + sheetsLocal.Count * (sheetWidth + sheetGap),
                    OriginY = baseOriginY
                };

                sheetsLocal.Add(sheet);

                DrawSheetOutline(sheet, sheetWidth, sheetHeight, modelSpace, layerSheet);

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
                    if (TryPlaceOnSheet(
                        sheetState,
                        inst,
                        partGap,
                        modelSpace,
                        anglesDeg,
                        layerParts,
                        placementsLocal,   // ✅ use local
                        ref placed,
                        totalInstances,
                        progress))
                    {
                        break;
                    }

                    // Could not fit on current sheet; add another
                    sheetState = NewSheet();
                }
            }

            // ✅ assign to out params once at end
            sheets = sheetsLocal;
            placements = placementsLocal;
        }

        /// <summary>
        /// Place the part in the best free rectangle (Best Short Side Fit, then Long Side Fit, then Area).
        /// Tries all allowed angles.
        /// </summary>
        private static bool TryPlaceOnSheet(
            SheetState sheet,
            PartDefinition part,
            double partGap,
            BlockRecord modelSpace,
            List<int> anglesDeg,
            object layerParts,
            List<PlacementRecord> placements,
            ref int placed,
            int totalInstances,
            LaserCutProgressForm progress)
        {
            if (sheet.FreeRects.Count == 0)
                return false;

            const double eps = 1e-9;

            int bestRectIndex = -1;
            int bestAngleDeg = 0;
            RotatedBounds bestBounds = default;

            double bestShortSideFit = double.MaxValue;
            double bestLongSideFit = double.MaxValue;
            double bestAreaFit = double.MaxValue;

            for (int i = 0; i < sheet.FreeRects.Count; i++)
            {
                var fr = sheet.FreeRects[i];

                for (int ai = 0; ai < anglesDeg.Count; ai++)
                {
                    int ang = anglesDeg[ai];
                    var rb = GetRotatedBounds(part, ang);

                    double usedW = rb.Width + partGap;
                    double usedH = rb.Height + partGap;

                    if (usedW > fr.Width + eps || usedH > fr.Height + eps)
                        continue;

                    double leftoverHoriz = fr.Width - usedW;
                    double leftoverVert = fr.Height - usedH;

                    double shortSideFit = Math.Min(leftoverHoriz, leftoverVert);
                    double longSideFit = Math.Max(leftoverHoriz, leftoverVert);
                    double areaFit = fr.Width * fr.Height - usedW * usedH;

                    if (shortSideFit < bestShortSideFit - eps ||
                        (Math.Abs(shortSideFit - bestShortSideFit) < eps && longSideFit < bestLongSideFit - eps) ||
                        (Math.Abs(shortSideFit - bestShortSideFit) < eps && Math.Abs(longSideFit - bestLongSideFit) < eps && areaFit < bestAreaFit - eps))
                    {
                        bestShortSideFit = shortSideFit;
                        bestLongSideFit = longSideFit;
                        bestAreaFit = areaFit;

                        bestRectIndex = i;
                        bestAngleDeg = ang;
                        bestBounds = rb;
                    }
                }
            }

            if (bestRectIndex < 0)
                return false;

            var chosen = sheet.FreeRects[bestRectIndex];

            // Place at the bottom-left corner of the chosen free rect.
            double usedX = chosen.X;
            double usedY = chosen.Y;

            // Actual part bbox min (inside used rect) leaves half gap on left/bottom.
            double partMinLocalX = usedX + partGap * 0.5;
            double partMinLocalY = usedY + partGap * 0.5;

            double worldMinX = sheet.OriginX + partMinLocalX;
            double worldMinY = sheet.OriginY + partMinLocalY;

            // Translate so that rotated bbox min becomes worldMin.
            double insertX = worldMinX - bestBounds.MinX;
            double insertY = worldMinY - bestBounds.MinY;

            double rotRad = bestAngleDeg * Math.PI / 180.0;

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

            // Record placement
            placements.Add(new PlacementRecord
            {
                SheetIndex = sheet.Index,
                BlockName = part.BlockName,
                AngleDeg = bestAngleDeg,
                LocalMinX = partMinLocalX,
                LocalMinY = partMinLocalY,
                BoundW = bestBounds.Width,
                BoundH = bestBounds.Height,
                InsertWorldX = insertX,
                InsertWorldY = insertY
            });

            // Update sheet stats
            sheet.PlacedCount++;
            sheet.UsedArea += bestBounds.Width * bestBounds.Height;

            // Subtract used rectangle from free space (MaxRects style)
            var usedRect = new FreeRect
            {
                X = usedX,
                Y = usedY,
                Width = bestBounds.Width + partGap,
                Height = bestBounds.Height + partGap
            };

            SubtractUsedRect(sheet, usedRect);

            placed++;
            progress.Step($"Placed {placed} of {totalInstances} plates...");

            return true;
        }

        private static bool Intersects(FreeRect a, FreeRect b)
        {
            return !(a.X + a.Width <= b.X ||
                     b.X + b.Width <= a.X ||
                     a.Y + a.Height <= b.Y ||
                     b.Y + b.Height <= a.Y);
        }

        /// <summary>
        /// MaxRects-style split: for every free rectangle that intersects the used area,
        /// split it into up to 4 rectangles (left/right/bottom/top).
        /// Then prune contained rects and merge adjacent.
        /// </summary>
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

                // Left strip
                if (used.X > fr.X + eps)
                {
                    newFree.Add(new FreeRect
                    {
                        X = fr.X,
                        Y = fr.Y,
                        Width = used.X - fr.X,
                        Height = fr.Height
                    });
                }

                // Right strip
                if (usedRight < frRight - eps)
                {
                    newFree.Add(new FreeRect
                    {
                        X = usedRight,
                        Y = fr.Y,
                        Width = frRight - usedRight,
                        Height = fr.Height
                    });
                }

                // Bottom strip
                if (used.Y > fr.Y + eps)
                {
                    newFree.Add(new FreeRect
                    {
                        X = fr.X,
                        Y = fr.Y,
                        Width = fr.Width,
                        Height = used.Y - fr.Y
                    });
                }

                // Top strip
                if (usedTop < frTop - eps)
                {
                    newFree.Add(new FreeRect
                    {
                        X = fr.X,
                        Y = usedTop,
                        Width = fr.Width,
                        Height = frTop - usedTop
                    });
                }
            }

            // Drop tiny rectangles
            sheet.FreeRects = newFree
                .Where(r => r.Width > minSize && r.Height > minSize)
                .ToList();

            PruneContained(sheet.FreeRects);
            MergeAdjacent(sheet.FreeRects);
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

                        // Vertical merge: same X/Width, touching Y edges
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

                        // Horizontal merge: same Y/Height, touching X edges
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
            }
            while (changed);
        }

        #endregion

        #region Layers + sheet labels + report

        private static void DrawSheetOutline(
            SheetState sheet,
            double sheetWidth,
            double sheetHeight,
            BlockRecord modelSpace,
            object layerSheet)
        {
            var bottom = new Line
            {
                StartPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0),
                EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0)
            };
            var right = new Line
            {
                StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0),
                EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0)
            };
            var top = new Line
            {
                StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0),
                EndPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0)
            };
            var left = new Line
            {
                StartPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0),
                EndPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0)
            };

            SetEntityLayer(bottom, layerSheet, "SHEET");
            SetEntityLayer(right, layerSheet, "SHEET");
            SetEntityLayer(top, layerSheet, "SHEET");
            SetEntityLayer(left, layerSheet, "SHEET");

            modelSpace.Entities.Add(bottom);
            modelSpace.Entities.Add(right);
            modelSpace.Entities.Add(top);
            modelSpace.Entities.Add(left);
        }

        private static void AddSheetLabel(
            BlockRecord modelSpace,
            SheetState sheet,
            double sheetWidth,
            double sheetHeight,
            double sheetMargin,
            double fillPercent,
            object layerLabels)
        {
            // Put label slightly above sheet
            double x = sheet.OriginX + sheetMargin;
            double y = sheet.OriginY + sheetHeight + sheetMargin;

            string text =
                $"Sheet {sheet.Index} | Parts: {sheet.PlacedCount} | Fill: {fillPercent:0.0}%";

            var mt = new MText
            {
                Value = text,
                InsertPoint = new XYZ(x, y, 0.0),
                Height = 25.0
            };

            SetEntityLayer(mt, layerLabels, "LABELS");
            modelSpace.Entities.Add(mt);
        }

        private static void WriteReportCsv(
            string path,
            double sheetWidth,
            double sheetHeight,
            double placementMargin,
            double partGap,
            List<SheetState> sheets,
            List<PlacementRecord> placements,
            double usableWidth,
            double usableHeight)
        {
            double usableArea = usableWidth * usableHeight;

            var sb = new StringBuilder();

            sb.AppendLine("LaserCutNestingReport");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "SheetWidth_mm,{0}", sheetWidth));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "SheetHeight_mm,{0}", sheetHeight));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "PlacementMargin_mm,{0}", placementMargin));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "PartGap_mm,{0}", partGap));
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "SheetsUsed,{0}", sheets.Count));
            sb.AppendLine();

            sb.AppendLine("SheetIndex,OriginX,OriginY,PlacedParts,UsedArea,FillPercent");
            foreach (var s in sheets)
            {
                double fill = usableArea > 1e-9 ? (s.UsedArea / usableArea) * 100.0 : 0.0;
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3},{4},{5}",
                    s.Index,
                    s.OriginX,
                    s.OriginY,
                    s.PlacedCount,
                    s.UsedArea,
                    fill.ToString("0.0", CultureInfo.InvariantCulture)));
            }

            sb.AppendLine();
            sb.AppendLine("Placements");
            sb.AppendLine("SheetIndex,BlockName,AngleDeg,LocalMinX,LocalMinY,BoundW,BoundH,InsertWorldX,InsertWorldY");
            foreach (var p in placements)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3},{4},{5},{6},{7},{8}",
                    p.SheetIndex,
                    EscapeCsv(p.BlockName),
                    p.AngleDeg,
                    p.LocalMinX,
                    p.LocalMinY,
                    p.BoundW,
                    p.BoundH,
                    p.InsertWorldX,
                    p.InsertWorldY));
            }

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
                return "";

            bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            if (!mustQuote)
                return value;

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        #endregion

        #region Existing helpers (extents + load blocks + qty parsing)

        private static void GetModelSpaceExtents(
            CadDocument doc,
            out double minX,
            out double minY,
            out double maxX,
            out double maxY)
        {
            var modelSpace = doc.BlockRecords["*Model_Space"];

            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            foreach (var ent in modelSpace.Entities)
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
                catch
                {
                    // ignore entities without bbox
                }
            }

            if (minX == double.MaxValue || maxX == double.MinValue)
            {
                minX = 0.0;
                minY = 0.0;
                maxX = 0.0;
                maxY = 0.0;
            }
        }

        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc)
        {
            var list = new List<PartDefinition>();

            foreach (var br in doc.BlockRecords)
            {
                if (string.IsNullOrEmpty(br.Name))
                    continue;

                if (br.Name.StartsWith("*", StringComparison.Ordinal))
                    continue;

                if (!br.Name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                int qty = ParseQuantityFromBlockName(br.Name);
                if (qty <= 0)
                    qty = 1;

                double minX = double.MaxValue;
                double minY = double.MaxValue;
                double maxX = double.MinValue;
                double maxY = double.MinValue;

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
                    catch
                    {
                        // ignore entities without bbox
                    }
                }

                if (minX == double.MaxValue || maxX == double.MinValue ||
                    minY == double.MaxValue || maxY == double.MinValue)
                {
                    continue;
                }

                double width = maxX - minX;
                double height = maxY - minY;

                if (width <= 0.0 || height <= 0.0)
                    continue;

                list.Add(new PartDefinition
                {
                    Block = br,
                    BlockName = br.Name,
                    MinX = minX,
                    MinY = minY,
                    MaxX = maxX,
                    MaxY = maxY,
                    Width = width,
                    Height = height,
                    Quantity = qty
                });
            }

            return list;
        }

        private static int ParseQuantityFromBlockName(string blockName)
        {
            if (string.IsNullOrEmpty(blockName))
                return 1;

            int idx = blockName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            if (idx < 0 || idx + 2 >= blockName.Length)
                return 1;

            string s = blockName.Substring(idx + 2);
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty) && qty > 0)
                return qty;

            return 1;
        }

        #endregion
    }

    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly TextBox _txtWidth;
        private readonly TextBox _txtHeight;

        private readonly ComboBox _cmbRotation;
        private readonly NumericUpDown _numStep;
        private readonly CheckBox _chkReport;

        private readonly Button _btnOk;
        private readonly Button _btnCancel;

        public double SheetWidthMm { get; private set; }
        public double SheetHeightMm { get; private set; }

        public RotationMode RotationMode { get; private set; } = RotationMode.AnyAngleStep;
        public int AnyAngleStepDeg { get; private set; } = 10;
        public bool WriteReportCsv { get; private set; } = false;

        public LaserCutOptionsForm()
        {
            Text = "Laser cut options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            ClientSize = new System.Drawing.Size(420, 230);

            var lblWidth = new Label
            {
                AutoSize = true,
                Text = "Sheet width (mm):",
                Location = new System.Drawing.Point(12, 18)
            };
            Controls.Add(lblWidth);

            _txtWidth = new TextBox
            {
                Location = new System.Drawing.Point(170, 14),
                Width = 200,
                Text = "3000"
            };
            Controls.Add(_txtWidth);

            var lblHeight = new Label
            {
                AutoSize = true,
                Text = "Sheet height (mm):",
                Location = new System.Drawing.Point(12, 52)
            };
            Controls.Add(lblHeight);

            _txtHeight = new TextBox
            {
                Location = new System.Drawing.Point(170, 48),
                Width = 200,
                Text = "1500"
            };
            Controls.Add(_txtHeight);

            var lblRot = new Label
            {
                AutoSize = true,
                Text = "Rotation mode:",
                Location = new System.Drawing.Point(12, 88)
            };
            Controls.Add(lblRot);

            _cmbRotation = new ComboBox
            {
                Location = new System.Drawing.Point(170, 84),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbRotation.Items.Add("0° only");
            _cmbRotation.Items.Add("0° / 90°");
            _cmbRotation.Items.Add("0° / 90° / 180° / 270°");
            _cmbRotation.Items.Add("Any angle (step)");
            _cmbRotation.SelectedIndex = 3;
            _cmbRotation.SelectedIndexChanged += (s, e) => UpdateStepEnabled();
            Controls.Add(_cmbRotation);

            var lblStep = new Label
            {
                AutoSize = true,
                Text = "Any-angle step (deg):",
                Location = new System.Drawing.Point(12, 122)
            };
            Controls.Add(lblStep);

            _numStep = new NumericUpDown
            {
                Location = new System.Drawing.Point(170, 118),
                Width = 80,
                Minimum = 1,
                Maximum = 90,
                Value = 10
            };
            Controls.Add(_numStep);

            _chkReport = new CheckBox
            {
                AutoSize = true,
                Text = "Write nest_report.csv",
                Location = new System.Drawing.Point(170, 152),
                Checked = false
            };
            Controls.Add(_chkReport);

            var note = new Label
            {
                AutoSize = false,
                Location = new System.Drawing.Point(12, 178),
                Size = new System.Drawing.Size(390, 28),
                Text = "Note: Part gap & margins are auto-calculated and never less than plate thickness."
            };
            Controls.Add(note);

            _btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.None,
                Location = new System.Drawing.Point(236, 205),
                Width = 75
            };
            _btnOk.Click += Ok_Click;
            Controls.Add(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(317, 205),
                Width = 75
            };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            UpdateStepEnabled();
        }

        private void UpdateStepEnabled()
        {
            bool anyAngle = _cmbRotation.SelectedIndex == 3;
            _numStep.Enabled = anyAngle;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(_txtWidth.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double w) || w <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet width (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtWidth.Focus();
                _txtWidth.SelectAll();
                return;
            }

            if (!double.TryParse(_txtHeight.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double h) || h <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet height (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHeight.Focus();
                _txtHeight.SelectAll();
                return;
            }

            SheetWidthMm = w;
            SheetHeightMm = h;

            switch (_cmbRotation.SelectedIndex)
            {
                case 0: RotationMode = RotationMode.Deg0Only; break;
                case 1: RotationMode = RotationMode.Deg0_90; break;
                case 2: RotationMode = RotationMode.Deg0_90_180_270; break;
                default: RotationMode = RotationMode.AnyAngleStep; break;
            }

            AnyAngleStepDeg = (int)_numStep.Value;
            WriteReportCsv = _chkReport.Checked;

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    internal sealed class LaserCutProgressForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _label;
        private readonly int _maximum;
        private int _current;

        public LaserCutProgressForm(int maximum)
        {
            if (maximum <= 0)
                maximum = 1;

            _maximum = maximum;

            Text = "Working...";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            ClientSize = new System.Drawing.Size(400, 90);

            _label = new Label
            {
                AutoSize = false,
                Text = "Preparing...",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Location = new System.Drawing.Point(12, 9),
                Size = new System.Drawing.Size(376, 20)
            };
            Controls.Add(_label);

            _progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(376, 20),
                Minimum = 0,
                Maximum = _maximum,
                Value = 0
            };
            Controls.Add(_progressBar);
        }

        public void Step(string statusText)
        {
            if (!IsHandleCreated)
                return;

            if (!string.IsNullOrEmpty(statusText))
                _label.Text = statusText;

            if (_current < _maximum)
            {
                _current++;
                _progressBar.Value = _current;
            }

            _progressBar.Refresh();
            _label.Refresh();
            Application.DoEvents();
        }
    }
}
