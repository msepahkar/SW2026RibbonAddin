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
using CSMath; // XYZ, BoundingBox, etc.

namespace SW2026RibbonAddin.Commands
{
    internal sealed class CombineDwgButton : IMehdiRibbonButton
    {
        public string Id => "CombineDwg";

        public string DisplayName => "Combine\nDWG";
        public string Tooltip => "Combine DWG exports from multiple jobs into per-thickness DWGs and a summary CSV.";
        public string Hint => "Combine DWG exports";

        public string SmallIconFile => "combine_dwg_20.png";
        public string LargeIconFile => "combine_dwg_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 2;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolder();
            if (string.IsNullOrEmpty(mainFolder))
                return;

            try
            {
                DwgBatchCombiner.Combine(mainFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Combine DWG failed:\r\n\r\n" + ex.Message,
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // Independent of active SW document
            return AddinContext.Enable;
        }

        private static string SelectMainFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the MAIN folder that contains job subfolders (each with parts.csv + DWGs)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }
    }

    internal static class DwgBatchCombiner
    {
        private static readonly Random _random = new Random();

        // ------------------------------------------------------------
        // COLOR PICKING (random, but always visible on black background)
        //
        // We pick random ACI colors from bright groups (no near-black colors).
        // We also build a queue by interleaving groups so two similar hues
        // are less likely to appear next to each other.
        // ------------------------------------------------------------

        // Bright ACI indices (hand-picked to have good luminance on black background)
        // and distinct hue families. Avoids very dark blues/purples, etc.
        //
        // Note: ACI 7 behaves as "white/black" depending on background; on black it shows as white.
        private static readonly byte[][] VisibleAciGroups =
        {
            new byte[] { 11, 241 },           // light reds / pinkish reds
            new byte[] { 20, 30 },            // oranges
            new byte[] { 40, 50 },            // yellows
            new byte[] { 60, 70 },            // yellow-green / lime
            new byte[] { 80, 90 },            // greens
            new byte[] { 100, 110 },          // green-cyans
            new byte[] { 120, 130 },          // cyans
            new byte[] { 140, 150, 161, 171 },// brighter blues (avoid pure dark ACI 5)
            new byte[] { 191, 201 },          // purples (brighter)
            new byte[] { 210, 220, 230 },     // magenta -> pink
            new byte[] { 7 }                  // white
        };

        private static readonly Queue<byte> _visibleAciQueue = new Queue<byte>();

        private static void Shuffle<T>(T[] array)
        {
            for (int i = array.Length - 1; i > 0; i--)
            {
                int j = _random.Next(i + 1);
                T tmp = array[i];
                array[i] = array[j];
                array[j] = tmp;
            }
        }

        private static void RefillVisibleColorQueue()
        {
            _visibleAciQueue.Clear();

            // Randomize group order each refill
            int groupCount = VisibleAciGroups.Length;
            int[] groupOrder = Enumerable.Range(0, groupCount).ToArray();
            Shuffle(groupOrder);

            // Shuffle colors inside each group and keep them in queues
            var perGroupQueues = new Queue<byte>[groupCount];
            for (int i = 0; i < groupCount; i++)
            {
                byte[] colors = (byte[])VisibleAciGroups[i].Clone();
                Shuffle(colors);
                perGroupQueues[i] = new Queue<byte>(colors);
            }

            // Interleave groups: take 1 color from each group in order, repeat,
            // until all groups run out. This reduces consecutive similar colors.
            bool addedAny;
            do
            {
                addedAny = false;

                foreach (int gi in groupOrder)
                {
                    if (perGroupQueues[gi].Count > 0)
                    {
                        _visibleAciQueue.Enqueue(perGroupQueues[gi].Dequeue());
                        addedAny = true;
                    }
                }
            }
            while (addedAny);
        }

        private static Color NextVisibleColor()
        {
            if (_visibleAciQueue.Count == 0)
                RefillVisibleColorQueue();

            return new Color(_visibleAciQueue.Dequeue());
        }

        private sealed class CombinedPart
        {
            public string FileName;
            public string FolderName;
            public string FullPath;
            public double ThicknessMm;
            public int Quantity;
        }

        public static void Combine(string mainFolder)
        {
            if (string.IsNullOrEmpty(mainFolder) || !Directory.Exists(mainFolder))
            {
                MessageBox.Show("The selected folder does not exist.", "Combine DWG",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string[] subFolders;
            try
            {
                subFolders = Directory.GetDirectories(mainFolder, "*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to enumerate subfolders:\r\n\r\n" + ex.Message,
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (subFolders.Length == 0)
            {
                MessageBox.Show("The selected folder does not contain any subfolders.",
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            // --- read all parts.csv and merge rows ---
            foreach (string sub in subFolders)
            {
                string csvPath = Path.Combine(sub, "parts.csv");
                if (!File.Exists(csvPath))
                    continue;

                string folderName = Path.GetFileName(sub);

                string[] lines;
                try
                {
                    lines = File.ReadAllLines(csvPath);
                }
                catch
                {
                    continue;
                }

                if (lines.Length <= 1)
                    continue;

                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] cols = line.Split(',');
                    if (cols.Length < 3)
                        continue;

                    string fileName = cols[0].Trim();
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    if (!double.TryParse(cols[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double tMm))
                        continue;

                    if (!int.TryParse(cols[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                        continue;

                    string key = MakeKey(fileName, tMm);

                    if (!combined.TryGetValue(key, out CombinedPart part))
                    {
                        part = new CombinedPart
                        {
                            FileName = fileName,
                            FolderName = folderName,
                            FullPath = Path.Combine(sub, fileName),
                            ThicknessMm = tMm,
                            Quantity = 0
                        };
                        combined.Add(key, part);
                    }

                    part.Quantity += qty;
                }
            }

            if (combined.Count == 0)
            {
                MessageBox.Show("No parts.csv files with data were found in any subfolder.",
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var list = new List<CombinedPart>(combined.Values);
            list.Sort((a, b) =>
            {
                int cmp = a.ThicknessMm.CompareTo(b.ThicknessMm);
                if (cmp != 0) return cmp;
                return string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase);
            });

            // --- write all_parts.csv in MAIN folder ---
            string allCsvPath = Path.Combine(mainFolder, "all_parts.csv");
            var outLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity,Folder"
            };

            foreach (var p in list)
            {
                outLines.Add(
                    string.Format(CultureInfo.InvariantCulture,
                        "{0},{1},{2},{3}",
                        p.FileName,
                        p.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture),
                        p.Quantity,
                        p.FolderName));
            }

            File.WriteAllLines(allCsvPath, outLines);

            // --- create per-thickness DWG files ---
            CreatePerThicknessDwgs(mainFolder, list);

            MessageBox.Show(
                "DWG combination finished.\r\n\r\n" +
                "Unique parts: " + list.Count + Environment.NewLine +
                "Summary CSV: " + allCsvPath,
                "Combine DWG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string MakeKey(string fileName, double thicknessMm)
        {
            return fileName.Trim().ToUpperInvariant() + "|" +
                   thicknessMm.ToString("0.###", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Block name for a plate: P_{sanitizedPartName}_Q{quantity}
        /// Quantity suffix is used later by the laser-cut nesting.
        /// </summary>
        private static string MakePlateBlockName(string fileName, int quantity)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Part";

            var sb = new StringBuilder();
            foreach (char c in baseName)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);
                else
                    sb.Append('_');
            }

            string safe = sb.Length > 0 ? sb.ToString() : "Part";
            int q = Math.Max(1, quantity);

            return string.Format(CultureInfo.InvariantCulture, "P_{0}_Q{1}", safe, q);
        }

        /// <summary>
        /// For each thickness, create thickness_XXX.dwg with all plates of that thickness
        /// laid out side-by-side:
        /// - bottoms aligned on Y = 0
        /// - two text lines beneath each plate (thickness + quantity)
        /// - horizontal spacing based on max(plate width, text width + padding) + margin
        /// </summary>
        private static void CreatePerThicknessDwgs(string mainFolder, List<CombinedPart> parts)
        {
            var groups = parts
                .GroupBy(p => p.ThicknessMm)
                .OrderBy(g => g.Key);

            foreach (var g in groups)
            {
                // New drawing => reset colors so you get a fresh random sequence per output DWG.
                RefillVisibleColorQueue();

                double thickness = g.Key;
                string thicknessText = thickness.ToString("0.###", CultureInfo.InvariantCulture);
                string fileSafeThickness = thicknessText.Replace('.', '_').Replace(',', '_');

                string outPath = Path.Combine(mainFolder, $"thickness_{fileSafeThickness}.dwg");

                var doc = new CadDocument();
                BlockRecord modelSpace = doc.BlockRecords["*Model_Space"];

                double cursorX = 0.0;
                const double marginX = 50.0;  // base margin between plates

                foreach (var part in g)
                {
                    if (!File.Exists(part.FullPath))
                        continue;

                    CadDocument srcDoc;
                    try
                    {
                        using (var reader = new DwgReader(part.FullPath))
                        {
                            srcDoc = reader.Read();
                        }
                    }
                    catch
                    {
                        continue;
                    }

                    BlockRecord srcModel;
                    try
                    {
                        srcModel = srcDoc.BlockRecords["*Model_Space"];
                    }
                    catch
                    {
                        continue;
                    }

                    // Create a block for this plate and copy all model space entities into it.
                    // Block name encodes the quantity for later laser nesting: P_..._Q{qty}
                    string blockName = MakePlateBlockName(part.FileName, part.Quantity);
                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    // ✅ Visible + distinct-ish color for black background (random from bright palette)
                    var blockColor = NextVisibleColor();

                    foreach (var ent in srcModel.Entities)
                    {
                        if (ent == null)
                            continue;

                        var cloned = ent.Clone() as Entity;
                        if (cloned == null)
                            continue;

                        cloned.Color = blockColor;
                        block.Entities.Add(cloned);
                    }

                    if (block.Entities.Count == 0)
                        continue;

                    // ---- get bounding box of block geometry (local coords) ----
                    double minX = double.MaxValue;
                    double maxX = double.MinValue;
                    double minY = double.MaxValue;

                    foreach (var ent in block.Entities)
                    {
                        try
                        {
                            var bb = ent.GetBoundingBox();   // BoundingBox is a struct, never null
                            XYZ bbMin = bb.Min;
                            XYZ bbMax = bb.Max;

                            if (bbMin.X < minX) minX = bbMin.X;
                            if (bbMax.X > maxX) maxX = bbMax.X;
                            if (bbMin.Y < minY) minY = bbMin.Y;
                        }
                        catch
                        {
                            // ignore entities that do not support bounding box
                        }
                    }

                    if (minX == double.MaxValue || maxX == double.MinValue)
                    {
                        // fallback if we could not compute a bbox
                        minX = 0.0;
                        maxX = 0.0;
                        minY = 0.0;
                    }

                    double blockWidth = maxX - minX;
                    if (blockWidth <= 0.0)
                        blockWidth = 1.0; // avoid zero-width issues

                    // ---- text under plate ----
                    double textHeight = 20.0;

                    // More comfortable vertical spacing:
                    double baselineY = 0.0;
                    double gapPlateToFirst = 8.0;     // plate bottom -> first line
                    double gapBetweenLines = 10.0;    // between line 1 and line 2

                    double textY1 = baselineY - textHeight - gapPlateToFirst;
                    double textY2 = textY1 - textHeight - gapBetweenLines;

                    string label1 = $"Plate: {thicknessText} mm";
                    string label2 = $"Qty: {part.Quantity}";

                    double textWidth1 = EstimateTextWidth(label1, textHeight);
                    double textWidth2 = EstimateTextWidth(label2, textHeight);
                    double maxTextWidth = Math.Max(textWidth1, textWidth2);

                    // Column width: if text is wider than plate, give extra horizontal
                    // room so that texts from neighboring plates do not overlap.
                    double columnWidth;
                    double extraTextSidePadding = textHeight; // padding each side when text controls width

                    if (maxTextWidth > blockWidth)
                    {
                        columnWidth = maxTextWidth + 2.0 * extraTextSidePadding;
                    }
                    else
                    {
                        columnWidth = blockWidth;
                    }

                    // Center of this column (also block + text center)
                    double columnCenterX = cursorX + columnWidth / 2.0;

                    // Align block bottom to Y = 0 and center it in the column.
                    double blockCenterLocalX = (minX + maxX) * 0.5;
                    double insertX = columnCenterX - blockCenterLocalX;
                    double insertY = -minY;

                    var insert = new Insert(block)
                    {
                        InsertPoint = new XYZ(insertX, insertY, 0.0),
                        XScale = 1.0,
                        YScale = 1.0,
                        ZScale = 1.0
                    };

                    doc.Entities.Add(insert);

                    // Center text on the same column center.
                    double plateCenterX = columnCenterX;

                    // Shift insertion points so that text is centered on the plate.
                    double text1InsertX = plateCenterX - textWidth1 / 2.0;
                    double text2InsertX = plateCenterX - textWidth2 / 2.0;

                    var text1 = new MText
                    {
                        Value = label1,
                        InsertPoint = new XYZ(text1InsertX, textY1, 0.0),
                        Height = textHeight
                    };

                    var text2 = new MText
                    {
                        Value = label2,
                        InsertPoint = new XYZ(text2InsertX, textY2, 0.0),
                        Height = textHeight
                    };

                    doc.Entities.Add(text1);
                    doc.Entities.Add(text2);

                    // Move cursor to the right of this plate, based on the column width
                    cursorX += columnWidth + marginX;
                }

                using (var writer = new DwgWriter(outPath, doc))
                {
                    writer.Write();
                }
            }
        }

        // Conservative factor so the estimated width is slightly larger than real text,
        // which helps guarantee that neighboring texts do not overlap.
        private const double TextWidthFactor = 1.0;

        private static double EstimateTextWidth(string text, double textHeight)
        {
            if (string.IsNullOrEmpty(text) || textHeight <= 0.0)
                return 0.0;

            // Very simple approximation: width ≈ characters * textHeight * factor
            return text.Length * textHeight * TextWidthFactor;
        }
    }
}
