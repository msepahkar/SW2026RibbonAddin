using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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

        private sealed class CombinedPart
        {
            public string FileName;
            public string FullPath;
            public double ThicknessMm;
            public int Quantity;

            // Identity protection:
            // If two different jobs have the same FileName and thickness, we only merge if the hash matches.
            public string HashHex; // full SHA256 hex or null if file missing/unreadable

            // Source folders (can be multiple jobs)
            public HashSet<string> SourceFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public string HashShort
            {
                get
                {
                    if (string.IsNullOrWhiteSpace(HashHex)) return null;
                    return HashHex.Length <= 8 ? HashHex : HashHex.Substring(0, 8);
                }
            }
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

            // Key strategy:
            // - If DWG exists -> FileName + Thickness + SHA256 => merge only identical geometry.
            // - If DWG missing/unreadable -> FileName + Thickness + FolderName => do NOT merge across jobs.
            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            int rowsRead = 0;
            int missingDwgCount = 0;

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

                    string fileName = (cols[0] ?? "").Trim();
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    if (!double.TryParse((cols[1] ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double tMm))
                        continue;

                    if (!int.TryParse((cols[2] ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                        continue;

                    rowsRead++;

                    string fullPath = Path.Combine(sub, fileName);

                    string hashHex = null;
                    if (File.Exists(fullPath))
                    {
                        hashHex = TryComputeSha256Hex(fullPath);
                    }
                    else
                    {
                        missingDwgCount++;
                    }

                    string key = MakeKey(fileName, tMm, hashHex, folderName);

                    if (!combined.TryGetValue(key, out CombinedPart part))
                    {
                        part = new CombinedPart
                        {
                            FileName = fileName,
                            FullPath = fullPath,
                            ThicknessMm = tMm,
                            Quantity = 0,
                            HashHex = hashHex
                        };
                        part.SourceFolders.Add(folderName);

                        combined.Add(key, part);
                    }
                    else
                    {
                        // If we already had a record and our stored FullPath is missing but this one exists, prefer existing file.
                        if (!File.Exists(part.FullPath) && File.Exists(fullPath))
                            part.FullPath = fullPath;

                        part.SourceFolders.Add(folderName);
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
                string folders = string.Join(";", p.SourceFolders.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

                outLines.Add(string.Format(CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3}",
                    EscapeCsv(p.FileName),
                    p.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture),
                    p.Quantity,
                    EscapeCsv(folders)));
            }

            File.WriteAllLines(allCsvPath, outLines);

            // --- create per-thickness DWG files ---
            CreatePerThicknessDwgs(mainFolder, list);

            MessageBox.Show(
                "DWG combination finished.\r\n\r\n" +
                "Rows read from CSVs: " + rowsRead + Environment.NewLine +
                "Unique parts: " + list.Count + Environment.NewLine +
                "Missing DWGs referenced by CSV: " + missingDwgCount + Environment.NewLine +
                "Summary CSV: " + allCsvPath,
                "Combine DWG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string MakeKey(string fileName, double thicknessMm, string hashHex, string folderName)
        {
            // If we can hash, use it to merge only identical DWGs.
            if (!string.IsNullOrWhiteSpace(hashHex))
            {
                return fileName.Trim().ToUpperInvariant() + "|" +
                       thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) + "|" +
                       hashHex.Trim().ToUpperInvariant();
            }

            // If DWG missing/unreadable, do not merge across folders (safer).
            return fileName.Trim().ToUpperInvariant() + "|" +
                   thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) + "|" +
                   (folderName ?? "").Trim().ToUpperInvariant();
        }

        private static string TryComputeSha256Hex(string filePath)
        {
            try
            {
                using (var sha = SHA256.Create())
                using (var fs = File.OpenRead(filePath))
                {
                    byte[] hash = sha.ComputeHash(fs);
                    return BytesToHex(hash);
                }
            }
            catch
            {
                return null;
            }
        }

        private static string BytesToHex(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;

            var sb = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++)
                sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));

            return sb.ToString();
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
                return "";

            bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            if (!mustQuote)
                return value;

            // Escape quotes by doubling them
            string escaped = value.Replace("\"", "\"\"");
            return "\"" + escaped + "\"";
        }

        /// <summary>
        /// Block name for a plate:
        /// P_{sanitizedPartName}[_H{hash8}]_Q{quantity}
        /// Quantity suffix is used later by the laser-cut nesting.
        /// </summary>
        private static string MakePlateBlockName(string fileName, string hashHex, int quantity)
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

            string h = null;
            if (!string.IsNullOrWhiteSpace(hashHex))
            {
                string hh = hashHex.Trim();
                h = hh.Length <= 8 ? hh : hh.Substring(0, 8);
            }

            // Keep _Q{qty} at the end so LaserCut parsing still works.
            if (!string.IsNullOrEmpty(h))
                return string.Format(CultureInfo.InvariantCulture, "P_{0}_H{1}_Q{2}", safe, h, q);

            return string.Format(CultureInfo.InvariantCulture, "P_{0}_Q{1}", safe, q);
        }

        private static bool BlockRecordExists(CadDocument doc, string blockName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(blockName))
                return false;

            try
            {
                // Indexer throws if not found
                var _ = doc.BlockRecords[blockName];
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string EnsureUniqueBlockName(CadDocument doc, string desiredName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(desiredName))
                return desiredName;

            if (!BlockRecordExists(doc, desiredName))
                return desiredName;

            // Insert suffix before _Q if possible, so qty stays parseable.
            int qIdx = desiredName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            string left = qIdx >= 0 ? desiredName.Substring(0, qIdx) : desiredName;
            string right = qIdx >= 0 ? desiredName.Substring(qIdx) : "";

            for (int i = 2; i < 9999; i++)
            {
                string candidate = left + "_D" + i.ToString(CultureInfo.InvariantCulture) + right;
                if (!BlockRecordExists(doc, candidate))
                    return candidate;
            }

            // Extreme fallback
            return left + "_D" + Guid.NewGuid().ToString("N").Substring(0, 6) + right;
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
                    // Block name encodes the quantity for later laser nesting: ..._Q{qty}
                    string blockName = MakePlateBlockName(part.FileName, part.HashHex, part.Quantity);
                    blockName = EnsureUniqueBlockName(doc, blockName);

                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    // Pick a random ACI color (1..255) for all entities in this block (kept as original behavior)
                    var blockColor = new Color((byte)_random.Next(1, 256));

                    foreach (var ent in srcModel.Entities)
                    {
                        if (ent == null)
                            continue;

                        var cloned = ent.Clone() as Entity;
                        if (cloned == null)
                            continue;

                        // Apply the random block color
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
