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

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

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

        // Bright ACI colors (good on black background)
        private static readonly byte[] _brightAci = new byte[] { 1, 2, 3, 4, 5, 6, 7 };

        // ✅ FIX: byte, not int
        private static byte _lastAci = 0;
        private static bool _hasLastAci = false;

        private sealed class CombinedPart
        {
            public string FileName;
            public string FolderName;
            public string FullPath;
            public double ThicknessMm;
            public int Quantity;

            public string MaterialRaw;
            public string MaterialTag;
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
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (subFolders.Length == 0)
            {
                MessageBox.Show("The selected folder does not contain any subfolders.",
                    "Combine DWG",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            foreach (string sub in subFolders)
            {
                string csvPath = Path.Combine(sub, "parts.csv");
                if (!File.Exists(csvPath))
                    continue;

                string folderName = Path.GetFileName(sub);

                string[] lines;
                try { lines = File.ReadAllLines(csvPath); }
                catch { continue; }

                if (lines.Length <= 1)
                    continue;

                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    List<string> cols = ParseCsvLine(line);
                    if (cols.Count < 3)
                        continue;

                    string fileName = (cols[0] ?? "").Trim();
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    if (!double.TryParse((cols[1] ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double tMm))
                        continue;

                    if (!int.TryParse((cols[2] ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                        continue;

                    string materialRaw = cols.Count >= 4 ? (cols[3] ?? "").Trim() : "UNKNOWN";
                    if (string.IsNullOrWhiteSpace(materialRaw))
                        materialRaw = "UNKNOWN";

                    string materialTag = SanitizeMaterialTag(materialRaw);

                    string key = MakeKey(fileName, tMm, materialTag);

                    if (!combined.TryGetValue(key, out CombinedPart part))
                    {
                        part = new CombinedPart
                        {
                            FileName = fileName,
                            FolderName = folderName,
                            FullPath = Path.Combine(sub, fileName),
                            ThicknessMm = tMm,
                            Quantity = 0,
                            MaterialRaw = materialRaw,
                            MaterialTag = materialTag
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

                cmp = string.Compare(a.MaterialTag, b.MaterialTag, StringComparison.OrdinalIgnoreCase);
                if (cmp != 0) return cmp;

                return string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase);
            });

            string allCsvPath = Path.Combine(mainFolder, "all_parts.csv");
            var outLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity,Folder,Material"
            };

            foreach (var p in list)
            {
                outLines.Add(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3},{4}",
                    EscapeCsv(p.FileName),
                    p.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture),
                    p.Quantity,
                    EscapeCsv(p.FolderName),
                    EscapeCsv(p.MaterialRaw)));
            }

            File.WriteAllLines(allCsvPath, outLines, Encoding.UTF8);

            CreatePerThicknessDwgs(mainFolder, list);

            MessageBox.Show(
                "DWG combination finished.\r\n\r\n" +
                "Unique parts: " + list.Count + Environment.NewLine +
                "Summary CSV: " + allCsvPath,
                "Combine DWG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string MakeKey(string fileName, double thicknessMm, string materialTag)
        {
            string fn = (fileName ?? "").Trim().ToUpperInvariant();
            string mat = (materialTag ?? "UNKNOWN").Trim().ToUpperInvariant();

            return fn + "|" +
                   thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) + "|" +
                   mat;
        }

        private static string MakePlateBlockName(string fileName, int quantity, string materialTag)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Part";

            var sb = new StringBuilder();
            foreach (char c in baseName)
            {
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            }

            string safePart = sb.Length > 0 ? sb.ToString() : "Part";
            string safeMat = SanitizeMaterialTag(materialTag);

            int q = Math.Max(1, quantity);

            return string.Format(
                CultureInfo.InvariantCulture,
                "P_{0}__MAT_{1}__Q{2}",
                safePart,
                safeMat,
                q);
        }

        private static string SanitizeMaterialTag(string materialRaw)
        {
            string s = (materialRaw ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s))
                s = "UNKNOWN";

            s = s.Replace('_', ' ');

            var sb = new StringBuilder(s.Length);
            foreach (char c in s.ToUpperInvariant())
            {
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            }

            string tag = sb.ToString();
            while (tag.Contains("__"))
                tag = tag.Replace("__", "_");

            tag = tag.Trim('_');
            if (string.IsNullOrWhiteSpace(tag))
                tag = "UNKNOWN";

            const int maxLen = 32;
            if (tag.Length > maxLen)
                tag = tag.Substring(0, maxLen);

            return tag;
        }

        private static byte PickBrightAci()
        {
            if (_brightAci.Length == 0)
                return 7;

            for (int tries = 0; tries < 10; tries++)
            {
                byte pick = _brightAci[_random.Next(0, _brightAci.Length)];
                if (!_hasLastAci || pick != _lastAci || _brightAci.Length == 1)
                {
                    _lastAci = pick;
                    _hasLastAci = true;
                    return pick;
                }
            }

            _lastAci = _brightAci[0];
            _hasLastAci = true;
            return _lastAci;
        }

        private static List<string> ParseCsvLine(string line)
        {
            var cols = new List<string>();
            if (line == null)
                return cols;

            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (inQuotes)
                {
                    if (c == '"')
                    {
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
                    if (c == ',')
                    {
                        cols.Add(sb.ToString());
                        sb.Clear();
                    }
                    else if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
            }

            cols.Add(sb.ToString());
            return cols;
        }

        private static string EscapeCsv(string value)
        {
            if (value == null) return "";
            bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            if (!mustQuote) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static void CreatePerThicknessDwgs(string mainFolder, List<CombinedPart> parts)
        {
            var groups = parts.GroupBy(p => p.ThicknessMm).OrderBy(g => g.Key);

            foreach (var g in groups)
            {
                double thickness = g.Key;
                string thicknessText = thickness.ToString("0.###", CultureInfo.InvariantCulture);
                string fileSafeThickness = thicknessText.Replace('.', '_').Replace(',', '_');

                string outPath = Path.Combine(mainFolder, $"thickness_{fileSafeThickness}.dwg");

                var doc = new CadDocument();
                BlockRecord modelSpace = doc.BlockRecords["*Model_Space"];

                double cursorX = 0.0;
                const double marginX = 50.0;

                foreach (var part in g)
                {
                    if (!File.Exists(part.FullPath))
                        continue;

                    CadDocument srcDoc;
                    try
                    {
                        using (var reader = new DwgReader(part.FullPath))
                            srcDoc = reader.Read();
                    }
                    catch
                    {
                        continue;
                    }

                    BlockRecord srcModel;
                    try { srcModel = srcDoc.BlockRecords["*Model_Space"]; }
                    catch { continue; }

                    string blockName = MakePlateBlockName(part.FileName, part.Quantity, part.MaterialTag);
                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    var blockColor = new Color(PickBrightAci());

                    foreach (var ent in srcModel.Entities)
                    {
                        if (ent == null) continue;

                        var cloned = ent.Clone() as Entity;
                        if (cloned == null) continue;

                        cloned.Color = blockColor;
                        block.Entities.Add(cloned);
                    }

                    if (block.Entities.Count == 0)
                        continue;

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
                        catch { }
                    }

                    if (minX == double.MaxValue || maxX == double.MinValue)
                    {
                        minX = 0.0;
                        maxX = 0.0;
                        minY = 0.0;
                    }

                    double blockWidth = maxX - minX;
                    if (blockWidth <= 0.0) blockWidth = 1.0;

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

                    double extraTextSidePadding = textHeight;
                    double columnWidth = maxTextWidth > blockWidth
                        ? maxTextWidth + 2.0 * extraTextSidePadding
                        : blockWidth;

                    double columnCenterX = cursorX + columnWidth / 2.0;

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
                    writer.Write();
            }
        }

        private const double TextWidthFactor = 1.0;

        private static double EstimateTextWidth(string text, double textHeight)
        {
            if (string.IsNullOrEmpty(text) || textHeight <= 0.0)
                return 0.0;

            return text.Length * textHeight * TextWidthFactor;
        }
    }
}
