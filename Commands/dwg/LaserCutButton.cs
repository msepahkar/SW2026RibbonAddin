using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
        public string Tooltip => "Nest a combined DWG into laser sheets (optimized rectangle-based nesting).";
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

            // Ask for sheet size
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                sheetWidth = dlg.SheetWidthMm;
                sheetHeight = dlg.SheetHeightMm;
            }

            try
            {
                // Always use the optimized nesting algorithm
                DwgLaserNester.Nest(dwgPath, sheetWidth, sheetHeight);
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

            // Combined DWGs are produced with decimal separators replaced by '_',
            // e.g. thickness_2_5.dwg for a 2.5 mm plate.
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

                        // Remove trailing "mm" if present.
                        int mmIdx = after.IndexOf("mm", StringComparison.OrdinalIgnoreCase);
                        if (mmIdx >= 0)
                        {
                            after = after.Substring(0, mmIdx).Trim();
                        }

                        if (string.IsNullOrWhiteSpace(after))
                            continue;

                        // Normalise decimal separator and parse using invariant culture.
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
                // Best-effort only; ignore and fall back to defaults.
            }

            return null;
        }

        /// <summary>
        /// Get plate thickness in mm from either the DWG file name or the MText labels.
        /// </summary>
        private static double? TryGetPlateThicknessMm(CadDocument doc, string sourceDwgPath)
        {
            var fromFileName = TryGetPlateThicknessFromFileName(sourceDwgPath);
            if (fromFileName.HasValue)
                return fromFileName;

            return TryGetPlateThicknessFromMText(doc);
        }

        /// <summary>
        /// Reads a combined DWG (one produced by CombineDwg) and nests all plate blocks
        /// P_*_Q{qty} onto as many sheets as required using an optimized rectangle-based algorithm.
        /// </summary>
        public static void Nest(string sourceDwgPath, double sheetWidth, double sheetHeight)
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
            const double sheetMargin = 10.0;    // visible border (mm)
            const double defaultPartGap = 5.0;  // minimum nominal gap between parts (mm)
            const double sheetGap = 50.0;       // distance between sheets (mm)

            // Determine plate thickness (mm) if possible and ensure the gap
            // between plates is never smaller than the plate thickness.
            double partGap = defaultPartGap;
            double? plateThicknessMm = TryGetPlateThicknessMm(doc, sourceDwgPath);
            if (plateThicknessMm.HasValue && plateThicknessMm.Value > partGap)
            {
                partGap = plateThicknessMm.Value;
            }

            // We never place plates closer than this to the sheet border.
            // (border + 2 * gap gives clearance so nothing crosses the border)
            double placementMargin = sheetMargin + 2 * partGap;

            // Validate that every part fits inside a sheet with the inner margin.
            double usableWidth = sheetWidth - 2 * placementMargin;
            double usableHeight = sheetHeight - 2 * placementMargin;

            foreach (var p in parts)
            {
                if (p.Width > usableWidth || p.Height > usableHeight)
                {
                    throw new InvalidOperationException(
                        $"Part '{p.BlockName}' ({p.Width:0.##} x {p.Height:0.##} mm) " +
                        $"does not fit inside sheet {sheetWidth:0.##} x {sheetHeight:0.##} mm.");
                }
            }

            // Build list of instances (copies) to place.
            var instances = new List<PartDefinition>();
            foreach (var def in parts)
            {
                for (int i = 0; i < def.Quantity; i++)
                    instances.Add(def);
            }

            // Optimized algorithm: place largest parts first (by area).
            instances.Sort((a, b) =>
            {
                double areaA = a.Width * a.Height;
                double areaB = b.Width * b.Height;
                return areaB.CompareTo(areaA);
            });

            // Compute extents of the original combined layout (model space)
            GetModelSpaceExtents(doc, out double origMinX, out double origMinY, out double origMaxX, out double origMaxY);

            var modelSpace = doc.BlockRecords["*Model_Space"];

            // Place sheets ABOVE the original combined layout so the original
            // plates remain visible at the bottom of the nested DWG.
            double baseSheetOriginY = origMaxY + 200.0; // gap above original
            double baseSheetOriginX = origMinX;

            var progress = new LaserCutProgressForm(totalInstances)
            {
                Text = "Laser cut nesting (optimized)"
            };

            int sheetCount;
            try
            {
                progress.Show();
                Application.DoEvents();

                sheetCount = NestFreeRectangles(
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
                    totalInstances);
            }
            finally
            {
                progress.Close();
            }

            // Write nested result to a new DWG next to the source.
            string dir = Path.GetDirectoryName(sourceDwgPath);
            string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath);
            string outPath = Path.Combine(dir ?? string.Empty, nameNoExt + "_nested_optimized.dwg");

            using (var writer = new DwgWriter(outPath, doc))
            {
                writer.Write();
            }

            MessageBox.Show(
                "Laser cut nesting finished.\r\n\r\n" +
                "Algorithm: Optimized (rectangle-based)" + Environment.NewLine +
                "Sheets used: " + sheetCount + Environment.NewLine +
                "Total parts: " + totalInstances + Environment.NewLine +
                "Output DWG: " + outPath,
                "Laser Cut",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        #region Optimized free-rectangles algorithm (improved heuristics)

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
            int totalInstances)
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
                DrawSheetOutline(sheet, sheetWidth, sheetHeight, modelSpace);

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
                    if (TryPlaceOnSheetFreeRects(
                        sheetState,
                        inst,
                        partGap,
                        modelSpace,
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

            return sheets.Count;
        }

        /// <summary>
        /// Try to place a part on a sheet using a MaxRects-style heuristic:
        /// Best Short-Side Fit, then Best Long-Side Fit as tiebreaker.
        /// </summary>
        private static bool TryPlaceOnSheetFreeRects(
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
            bool bestRotated = false;
            double bestShortSideFit = double.MaxValue;
            double bestLongSideFit = double.MaxValue;

            // Try all free rectangles, both orientations (0° and 90°)
            for (int i = 0; i < sheet.FreeRects.Count; i++)
            {
                var fr = sheet.FreeRects[i];

                EvaluateOrientation(fr, part.Width, part.Height, false, i);
                EvaluateOrientation(fr, part.Height, part.Width, true, i);
            }

            if (bestIndex < 0)
                return false;

            var chosenRect = sheet.FreeRects[bestIndex];
            bool rotated = bestRotated;

            double w = rotated ? part.Height : part.Width;
            double h = rotated ? part.Width : part.Height;
            double placeW = w + partGap;
            double placeH = h + partGap;

            // Desired minimum corner of the part's bounding box (local to sheet)
            double desiredMinLocalX = chosenRect.X + partGap * 0.5;
            double desiredMinLocalY = chosenRect.Y + partGap * 0.5;

            double insertXWorld;
            double insertYWorld;
            double rotation;

            if (!rotated)
            {
                // No rotation: align bounding min directly
                double worldMinX = sheet.OriginX + desiredMinLocalX;
                double worldMinY = sheet.OriginY + desiredMinLocalY;

                insertXWorld = worldMinX - part.MinX;
                insertYWorld = worldMinY - part.MinY;
                rotation = 0.0;
            }
            else
            {
                // 90° rotation around insert point:
                // rotate (x,y) -> (-y, x).
                // Bounding min after rotation, about origin: (-MaxY, MinX).
                // After translation (Tx,Ty) => (Tx - MaxY, Ty + MinX).
                // We want these equal to desired world min coords.
                double worldMinX = sheet.OriginX + desiredMinLocalX;
                double worldMinY = sheet.OriginY + desiredMinLocalY;

                double maxYlocal = part.MaxY;

                insertXWorld = worldMinX + maxYlocal;   // Tx = worldMinX + MaxY
                insertYWorld = worldMinY - part.MinX;   // Ty = worldMinY - MinX
                rotation = Math.PI / 2.0;
            }

            var insert = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insertXWorld, insertYWorld, 0.0),
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0,
                Rotation = rotation
            };
            modelSpace.Entities.Add(insert);

            // Split the used free rectangle into right and top remainders
            SplitFreeRect(sheet, bestIndex, chosenRect, placeW, placeH);

            placed++;
            progress.Step($"Placed {placed} of {totalInstances} plates...");

            return true;

            void EvaluateOrientation(FreeRect fr, double wCandidate, double hCandidate, bool rotatedCandidate, int rectIndex)
            {
                double placeWCandidate = wCandidate + partGap;
                double placeHCandidate = hCandidate + partGap;

                if (placeWCandidate > fr.Width || placeHCandidate > fr.Height)
                    return;

                double leftoverHoriz = fr.Width - placeWCandidate;
                double leftoverVert = fr.Height - placeHCandidate;
                double shortSideFit = Math.Min(leftoverHoriz, leftoverVert);
                double longSideFit = Math.Max(leftoverHoriz, leftoverVert);

                if (shortSideFit < bestShortSideFit ||
                    (Math.Abs(shortSideFit - bestShortSideFit) < 1e-9 && longSideFit < bestLongSideFit))
                {
                    bestShortSideFit = shortSideFit;
                    bestLongSideFit = longSideFit;
                    bestIndex = rectIndex;
                    bestRotated = rotatedCandidate;
                }
            }
        }

        private static void SplitFreeRect(
            SheetState sheet,
            int rectIndex,
            FreeRect usedRect,
            double usedWidth,
            double usedHeight)
        {
            sheet.FreeRects.RemoveAt(rectIndex);

            const double minSize = 1.0;

            // Right remainder
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

            // Top remainder
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

        #endregion

        #region Helpers

        private static void DrawSheetOutline(
            SheetState sheet,
            double sheetWidth,
            double sheetHeight,
            BlockRecord modelSpace)
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

            modelSpace.Entities.Add(bottom);
            modelSpace.Entities.Add(right);
            modelSpace.Entities.Add(top);
            modelSpace.Entities.Add(left);
        }

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
        private readonly Button _btnOk;
        private readonly Button _btnCancel;

        public double SheetWidthMm { get; private set; }
        public double SheetHeightMm { get; private set; }

        public LaserCutOptionsForm()
        {
            Text = "Laser cut sheet size";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            ClientSize = new System.Drawing.Size(320, 140);

            var lblWidth = new Label
            {
                AutoSize = true,
                Text = "Sheet width (mm):",
                Location = new System.Drawing.Point(12, 20)
            };
            Controls.Add(lblWidth);

            _txtWidth = new TextBox
            {
                Location = new System.Drawing.Point(150, 16),
                Width = 140,
                Text = "3000"
            };
            Controls.Add(_txtWidth);

            var lblHeight = new Label
            {
                AutoSize = true,
                Text = "Sheet height (mm):",
                Location = new System.Drawing.Point(12, 55)
            };
            Controls.Add(lblHeight);

            _txtHeight = new TextBox
            {
                Location = new System.Drawing.Point(150, 51),
                Width = 140,
                Text = "1500"
            };
            Controls.Add(_txtHeight);

            _btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.None,
                Location = new System.Drawing.Point(134, 95),
                Width = 75
            };
            _btnOk.Click += Ok_Click;
            Controls.Add(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(215, 95),
                Width = 75
            };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(_txtWidth.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double w) ||
                w <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet width (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtWidth.Focus();
                _txtWidth.SelectAll();
                return;
            }

            if (!double.TryParse(_txtHeight.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double h) ||
                h <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet height (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHeight.Focus();
                _txtHeight.SelectAll();
                return;
            }

            SheetWidthMm = w;
            SheetHeightMm = h;

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
