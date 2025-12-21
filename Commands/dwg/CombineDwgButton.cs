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
                DwgBatchCombiner.Combine(mainFolder, showUi: true);
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

        // AutoCAD ACI indices that are very visible on black background
        private static readonly byte[] VisibleAciColors = { 1, 2, 3, 4, 5, 6, 7 };

        private sealed class CombinedPart
        {
            public string FileName;
            public string FolderName;
            public string FullPath;
            public double ThicknessMm;
            public int Quantity;
            public string MaterialExact; // EXACT string from SolidWorks (no normalization)
        }

        internal sealed class CombineRunResult
        {
            public int UniqueParts;
            public string AllCsvPath;
            public List<string> ThicknessDwgs = new List<string>();
        }

        public static CombineRunResult Combine(string mainFolder, bool showUi = true)
        {
            if (string.IsNullOrEmpty(mainFolder) || !Directory.Exists(mainFolder))
            {
                if (showUi)
                {
                    MessageBox.Show("The selected folder does not exist.", "Combine DWG",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return new CombineRunResult();
            }

            string[] subFolders;
            try
            {
                subFolders = Directory.GetDirectories(mainFolder, "*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                if (showUi)
                {
                    MessageBox.Show("Failed to enumerate subfolders:\r\n\r\n" + ex.Message,
                        "Combine DWG",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                return new CombineRunResult();
            }

            if (subFolders.Length == 0)
            {
                if (showUi)
                {
                    MessageBox.Show("The selected folder does not contain any subfolders.",
                        "Combine DWG",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return new CombineRunResult();
            }

            // Key includes FILE + THICKNESS + MATERIAL (exact string)
            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            // --- read all parts.csv and merge rows ---
            foreach (string sub in subFolders)
            {
                string csvPath = Path.Combine(sub, "parts.csv");
                if (!File.Exists(csvPath))
                    continue;

                string folderName = Path.GetFileName(sub);

                foreach (var row in ReadPartsCsv(csvPath))
                {
                    if (row == null)
                        continue;

                    string fileName = row.FileName;
                    double tMm = row.ThicknessMm;
                    int qty = row.Quantity;
                    string material = MaterialNameCodec.Normalize(row.MaterialExact);

                    if (string.IsNullOrWhiteSpace(fileName) || tMm <= 0 || qty <= 0)
                        continue;

                    string key = MakeKey(fileName, tMm, material);

                    if (!combined.TryGetValue(key, out CombinedPart part))
                    {
                        part = new CombinedPart
                        {
                            FileName = fileName,
                            FolderName = folderName,
                            FullPath = Path.Combine(sub, fileName),
                            ThicknessMm = tMm,
                            Quantity = 0,
                            MaterialExact = material
                        };
                        combined.Add(key, part);
                    }

                    part.Quantity += qty;
                }
            }

            if (combined.Count == 0)
            {
                if (showUi)
                {
                    MessageBox.Show("No parts.csv files with data were found in any subfolder.",
                        "Combine DWG",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return new CombineRunResult();
            }

            var list = new List<CombinedPart>(combined.Values);
            list.Sort((a, b) =>
            {
                int cmp = a.ThicknessMm.CompareTo(b.ThicknessMm);
                if (cmp != 0) return cmp;

                cmp = string.Compare(a.MaterialExact, b.MaterialExact, StringComparison.OrdinalIgnoreCase);
                if (cmp != 0) return cmp;

                return string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase);
            });

            // --- write all_parts.csv in MAIN folder ---
            string allCsvPath = Path.Combine(mainFolder, "all_parts.csv");
            var outLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity,Material,Folder"
            };

            foreach (var p in list)
            {
                outLines.Add(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3},{4}",
                    CsvCell(p.FileName),
                    p.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture),
                    p.Quantity,
                    CsvCell(p.MaterialExact),
                    CsvCell(p.FolderName)));
            }

            File.WriteAllLines(allCsvPath, outLines, Encoding.UTF8);

            // --- create per-thickness DWG files ---
            var thicknessDwgs = new List<string>();
            CreatePerThicknessDwgs(mainFolder, list, thicknessDwgs);

            if (showUi)
            {
                MessageBox.Show(
                    "DWG combination finished.\r\n\r\n" +
                    "Unique parts: " + list.Count + Environment.NewLine +
                    "Summary CSV: " + allCsvPath,
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            return new CombineRunResult
            {
                UniqueParts = list.Count,
                AllCsvPath = allCsvPath,
                ThicknessDwgs = thicknessDwgs
            };
        }

        private sealed class PartsCsvRow
        {
            public string FileName;
            public double ThicknessMm;
            public int Quantity;
            public string MaterialExact;
        }

        private static IEnumerable<PartsCsvRow> ReadPartsCsv(string csvPath)
        {
            string[] lines;
            try
            {
                lines = File.ReadAllLines(csvPath);
            }
            catch
            {
                yield break;
            }

            if (lines == null || lines.Length <= 1)
                yield break;

            // Find header
            int headerIndex = -1;
            for (int i = 0; i < lines.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(lines[i]))
                {
                    headerIndex = i;
                    break;
                }
            }

            if (headerIndex < 0 || headerIndex >= lines.Length - 1)
                yield break;

            var header = ParseCsvLine(lines[headerIndex]);
            int idxFile = FindHeaderIndex(header, "FileName", "Filename");
            int idxThk = FindHeaderIndex(header, "PlateThickness_mm", "Thickness", "PlateThickness");
            int idxQty = FindHeaderIndex(header, "Quantity", "Qty");
            int idxMat = FindHeaderIndex(header, "Material");

            // Backwards compatible: if old format (3 cols), assume:
            // 0=File, 1=Thickness, 2=Qty
            bool old3Col = (idxFile < 0 || idxThk < 0 || idxQty < 0);

            for (int i = headerIndex + 1; i < lines.Length; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var cols = ParseCsvLine(line);
                if (cols.Count < 3)
                    continue;

                string fileName;
                string thkStr;
                string qtyStr;
                string matStr = "UNKNOWN";

                if (old3Col)
                {
                    fileName = SafeGet(cols, 0);
                    thkStr = SafeGet(cols, 1);
                    qtyStr = SafeGet(cols, 2);

                    if (cols.Count >= 4)
                        matStr = SafeGet(cols, 3);
                }
                else
                {
                    fileName = SafeGet(cols, idxFile);
                    thkStr = SafeGet(cols, idxThk);
                    qtyStr = SafeGet(cols, idxQty);

                    if (idxMat >= 0)
                        matStr = SafeGet(cols, idxMat);
                }

                if (string.IsNullOrWhiteSpace(fileName))
                    continue;

                if (!double.TryParse((thkStr ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double tMm))
                    continue;

                if (!int.TryParse((qtyStr ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                    continue;

                yield return new PartsCsvRow
                {
                    FileName = fileName.Trim(),
                    ThicknessMm = tMm,
                    Quantity = qty,
                    MaterialExact = MaterialNameCodec.Normalize(matStr)
                };
            }
        }

        private static int FindHeaderIndex(List<string> header, params string[] candidates)
        {
            if (header == null || header.Count == 0)
                return -1;

            for (int i = 0; i < header.Count; i++)
            {
                string h = NormalizeHeader(header[i]);
                foreach (var c in candidates)
                {
                    if (h == NormalizeHeader(c))
                        return i;
                }
            }

            return -1;
        }

        private static string NormalizeHeader(string s)
        {
            s = (s ?? "").Trim();
            s = s.Replace(" ", "").Replace("-", "").Replace("_", "");
            return s.ToUpperInvariant();
        }

        private static string SafeGet(List<string> cols, int idx)
        {
            if (cols == null || idx < 0 || idx >= cols.Count)
                return "";
            return cols[idx] ?? "";
        }

        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            if (line == null)
                return result;

            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        // Escaped quote ""
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        result.Add(sb.ToString());
                        sb.Clear();
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
            }

            result.Add(sb.ToString());
            return result;
        }

        private static string MakeKey(string fileName, double thicknessMm, string materialExact)
        {
            return (fileName ?? "").Trim().ToUpperInvariant() + "|" +
                   thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) + "|" +
                   (materialExact ?? "").Trim().ToUpperInvariant();
        }

        /// <summary>
        /// Block name encodes:
        /// - part name (sanitized)
        /// - exact material string (encoded as Base64Url)
        /// - quantity suffix: _Q{qty}
        /// </summary>
        private static string MakePlateBlockName(string fileName, string materialExact, int quantity)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Part";

            // sanitize for block name readability (material stays reversible in token)
            var sb = new StringBuilder();
            foreach (char c in baseName)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);
                else
                    sb.Append('_');
            }

            string safe = sb.Length > 0 ? sb.ToString() : "Part";
            if (safe.Length > 80) safe = safe.Substring(0, 80);

            int q = Math.Max(1, quantity);
            string matToken = MaterialNameCodec.Encode(materialExact);

            return string.Format(
                CultureInfo.InvariantCulture,
                "P_{0}{1}{2}{3}_Q{4}",
                safe,
                MaterialNameCodec.BlockTokenPrefix,
                matToken,
                MaterialNameCodec.BlockTokenSuffix,
                q);
        }

        private static byte PickVisibleAciColor(byte? avoid = null)
        {
            for (int tries = 0; tries < 16; tries++)
            {
                byte c = VisibleAciColors[_random.Next(VisibleAciColors.Length)];
                if (!avoid.HasValue || c != avoid.Value)
                    return c;
            }

            return VisibleAciColors[0];
        }

        /// <summary>
        /// For each thickness, create thickness_XXX.dwg with all plates of that thickness
        /// laid out side-by-side:
        /// - bottoms aligned on Y = 0
        /// - two text lines beneath each plate
        /// </summary>
        private static void CreatePerThicknessDwgs(string mainFolder, List<CombinedPart> parts, List<string> outPaths)
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
                const double marginX = 50.0; // base margin between plates

                byte? lastColor = null;

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

                    // Block name includes EXACT material (encoded) + quantity
                    string baseBlockName = MakePlateBlockName(part.FileName, part.MaterialExact, part.Quantity);

                    // Ensure block name uniqueness in destination doc
                    BlockRecord block = null;
                    string blockName = baseBlockName;
                    for (int attempt = 0; attempt < 1000; attempt++)
                    {
                        try
                        {
                            block = new BlockRecord(blockName);
                            doc.BlockRecords.Add(block);
                            break;
                        }
                        catch
                        {
                            blockName = baseBlockName + "_" + (attempt + 1).ToString(CultureInfo.InvariantCulture);
                            block = null;
                        }
                    }

                    if (block == null)
                        continue;

                    // Pick a visible random ACI color (safe on black)
                    byte aci = PickVisibleAciColor(lastColor);
                    lastColor = aci;

                    var blockColor = new Color(aci);

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
                            var bb = ent.GetBoundingBox();
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
                        minX = 0.0;
                        maxX = 0.0;
                        minY = 0.0;
                    }

                    double blockWidth = maxX - minX;
                    if (blockWidth <= 0.0)
                        blockWidth = 1.0;

                    // ---- text under plate ----
                    double textHeight = 20.0;

                    double baselineY = 0.0;
                    double gapPlateToFirst = 8.0;
                    double gapBetweenLines = 10.0;

                    double textY1 = baselineY - textHeight - gapPlateToFirst;
                    double textY2 = textY1 - textHeight - gapBetweenLines;

                    string label1 = $"Plate: {thicknessText} mm";
                    string label2 = $"Qty: {part.Quantity}";

                    double textWidth1 = EstimateTextWidth(label1, textHeight);
                    double textWidth2 = EstimateTextWidth(label2, textHeight);
                    double maxTextWidth = Math.Max(textWidth1, textWidth2);

                    double extraTextSidePadding = textHeight; // generous side padding
                    double columnWidth = (maxTextWidth > blockWidth)
                        ? (maxTextWidth + 2.0 * extraTextSidePadding)
                        : blockWidth;

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

                    modelSpace.Entities.Add(insert);

                    // Center text under plate
                    double plateCenterX = columnCenterX;

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

                    modelSpace.Entities.Add(text1);
                    modelSpace.Entities.Add(text2);

                    cursorX += columnWidth + marginX;
                }

                using (var writer = new DwgWriter(outPath, doc))
                {
                    writer.Write();
                }

                outPaths.Add(outPath);
            }
        }

        // Conservative on purpose so texts won't overlap.
        private const double TextWidthFactor = 1.0;

        private static double EstimateTextWidth(string text, double textHeight)
        {
            if (string.IsNullOrEmpty(text) || textHeight <= 0.0)
                return 0.0;

            return text.Length * textHeight * TextWidthFactor;
        }

        private static string CsvCell(string s)
        {
            if (s == null) return "";
            s = s.Trim();

            bool needsQuotes =
                s.Contains(",") ||
                s.Contains("\"") ||
                s.Contains("\r") ||
                s.Contains("\n");

            if (!needsQuotes)
                return s;

            s = s.Replace("\"", "\"\"");
            return "\"" + s + "\"";
        }
    }

    /// <summary>
    /// Shared codec:
    /// - Combiner writes exact SolidWorks material into block name in a reversible safe token
    /// - Nester reads it back exactly (no normalization / no translation)
    /// </summary>
    internal static class MaterialNameCodec
    {
        public const string BlockTokenPrefix = "__MATB64_";
        public const string BlockTokenSuffix = "__";

        public static string Normalize(string material)
        {
            material = (material ?? "").Trim();
            return string.IsNullOrWhiteSpace(material) ? "UNKNOWN" : material;
        }

        // Base64Url (safe chars for DWG block names)
        public static string Encode(string material)
        {
            material = Normalize(material);

            byte[] bytes = Encoding.UTF8.GetBytes(material);
            string b64 = Convert.ToBase64String(bytes);

            // base64url: + -> -, / -> _, trim '='
            b64 = b64.TrimEnd('=').Replace('+', '-').Replace('/', '_');
            return string.IsNullOrWhiteSpace(b64) ? "VVNLTk9XTg" : b64; // fallback token
        }

        public static string Decode(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
                return "UNKNOWN";

            try
            {
                string s = token.Replace('-', '+').Replace('_', '/');
                int pad = s.Length % 4;
                if (pad != 0)
                    s = s + new string('=', 4 - pad);

                byte[] bytes = Convert.FromBase64String(s);
                return Normalize(Encoding.UTF8.GetString(bytes));
            }
            catch
            {
                return "UNKNOWN";
            }
        }

        public static bool TryExtractFromBlockName(string blockName, out string materialExact)
        {
            materialExact = "UNKNOWN";
            if (string.IsNullOrWhiteSpace(blockName))
                return false;

            int p = blockName.IndexOf(BlockTokenPrefix, StringComparison.Ordinal);
            if (p >= 0)
            {
                int start = p + BlockTokenPrefix.Length;
                int end = blockName.IndexOf(BlockTokenSuffix, start, StringComparison.Ordinal);
                if (end > start)
                {
                    string token = blockName.Substring(start, end - start);
                    materialExact = Decode(token);
                    return true;
                }
            }

            // Back-compat: old non-b64 token "__MAT_xxx__"
            const string oldPrefix = "__MAT_";
            int p2 = blockName.IndexOf(oldPrefix, StringComparison.Ordinal);
            if (p2 >= 0)
            {
                int start = p2 + oldPrefix.Length;
                int end = blockName.IndexOf("__", start, StringComparison.Ordinal);
                if (end > start)
                {
                    string token = blockName.Substring(start, end - start);
                    materialExact = Normalize(token.Replace('_', ' '));
                    return true;
                }
            }

            return false;
        }

        public static string MakeSafeFileToken(string materialExact, int maxLen = 40)
        {
            materialExact = Normalize(materialExact);

            var sb = new StringBuilder(materialExact.Length);
            foreach (char c in materialExact)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);
                else if (char.IsWhiteSpace(c))
                    sb.Append('_');
                else
                    sb.Append('_');
            }

            string safe = sb.ToString();
            while (safe.Contains("__"))
                safe = safe.Replace("__", "_");

            safe = safe.Trim('_');
            if (string.IsNullOrWhiteSpace(safe))
                safe = "UNKNOWN";

            if (safe.Length > maxLen)
                safe = safe.Substring(0, maxLen).Trim('_');

            return safe;
        }
    }
}
