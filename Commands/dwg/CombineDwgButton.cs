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
using AcadColor = ACadSharp.Color;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class CombineDwgButton : IMehdiRibbonButton
    {
        public string Id => "CombineDwg";

        public string DisplayName => "Combine\nDWG";
        public string Tooltip => "Combine job DWGs into per-thickness DWGs (now normalizes MaterialType).";
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
                MessageBox.Show("Combine DWG failed:\r\n\r\n" + ex.Message, "Combine DWG",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        private static string SelectMainFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select MAIN folder (contains job subfolders with parts.csv + DWGs)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }
    }

    // ✅ Shared normalizer used by Combine + LaserCut
    internal static class MaterialTypeNormalizer
    {
        public const string TYPE_STEEL = "STEEL";
        public const string TYPE_ALUMINUM = "ALUMINUM";
        public const string TYPE_STAINLESS = "STAINLESS";
        public const string TYPE_OTHER = "OTHER";

        public static string NormalizeToType(string rawOrTag)
        {
            string s = (rawOrTag ?? "").Trim();
            if (string.IsNullOrEmpty(s))
                return TYPE_OTHER;

            // strip long SW material strings: take last segment after :: or \ or /
            if (s.Contains("::"))
            {
                var parts = s.Split(new[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                s = parts.Length > 0 ? parts[parts.Length - 1] : s;
            }
            s = s.Replace('\\', '/');
            if (s.Contains("/"))
            {
                var parts = s.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                s = parts.Length > 0 ? parts[parts.Length - 1] : s;
            }

            string u = s.Trim().ToUpperInvariant();

            // quick keywords
            if (u.Contains("STAINLESS") || u.StartsWith("SS") || u.Contains(" AISI 304") || u.Contains("304") || u.Contains("316"))
                return TYPE_STAINLESS;

            if (u.Contains("ALUMIN") || u.StartsWith("ALU") || u.Contains("AL-") || u.Contains("6061") || u.Contains("5083") || u.Contains("7075"))
                return TYPE_ALUMINUM;

            // steel patterns (common in EU/IR shops)
            if (u.Contains("STEEL") || u.Contains("ST37") || u.Contains("ST-37") || u.Contains("S235") || u.Contains("S355") || u.Contains("A36") || u.Contains("CK") || u.Contains("C45"))
                return TYPE_STEEL;

            // fallback
            return TYPE_OTHER;
        }

        public static string MakeSafeTag(string s)
        {
            s = (s ?? "").Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(s)) s = TYPE_OTHER;

            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c)) sb.Append(c);
                else sb.Append('_');
            }

            string r = sb.ToString().Trim('_');
            return string.IsNullOrEmpty(r) ? TYPE_OTHER : r;
        }

        public static string TryParseMaterialFromFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            const string token = "__MAT_";
            int idx = fileName.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return null;

            int start = idx + token.Length;
            if (start >= fileName.Length) return null;

            int end = fileName.IndexOf("__", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0) end = fileName.Length;

            string tag = fileName.Substring(start, end - start);
            return string.IsNullOrWhiteSpace(tag) ? null : tag.Trim();
        }
    }

    internal static class DwgBatchCombiner
    {
        private static readonly Random _rnd = new Random();

        private sealed class CombinedPart
        {
            public string FileName;
            public string FolderName;
            public string FullPath;

            public double ThicknessMm;
            public int Quantity;

            public string MaterialRaw;
            public string MaterialType; // normalized: STEEL/ALUMINUM/STAINLESS/OTHER
            public string MaterialTag;  // safe: STEEL, ALUMINUM, ...
        }

        public static void Combine(string mainFolder, bool showUi)
        {
            if (string.IsNullOrWhiteSpace(mainFolder) || !Directory.Exists(mainFolder))
            {
                if (showUi) MessageBox.Show("Folder does not exist.", "Combine DWG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var subFolders = Directory.GetDirectories(mainFolder, "*", SearchOption.TopDirectoryOnly);
            if (subFolders.Length == 0)
            {
                if (showUi) MessageBox.Show("No job subfolders found.", "Combine DWG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // key includes material type (normalized)
            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            foreach (string sub in subFolders)
            {
                string folderName = Path.GetFileName(sub);
                if (string.IsNullOrWhiteSpace(folderName)) folderName = "Job";

                string csvPath = Path.Combine(sub, "parts.csv");
                if (!File.Exists(csvPath))
                    continue;

                string[] lines;
                try { lines = File.ReadAllLines(csvPath); }
                catch { continue; }

                if (lines.Length < 2)
                    continue;

                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var cols = SplitCsvSimple(line);
                    if (cols.Length < 3)
                        continue;

                    string fileName = cols[0].Trim().Trim('"');
                    if (string.IsNullOrWhiteSpace(fileName))
                        continue;

                    if (!double.TryParse(cols[1].Trim().Trim('"'), NumberStyles.Float, CultureInfo.InvariantCulture, out double thicknessMm))
                        continue;

                    if (!int.TryParse(cols[2].Trim().Trim('"'), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                        qty = 1;

                    // raw from CSV col4 if exists, else try parse from filename __MAT_
                    string materialRaw = cols.Length >= 4 ? cols[3].Trim().Trim('"') : null;
                    if (string.IsNullOrWhiteSpace(materialRaw))
                        materialRaw = MaterialTypeNormalizer.TryParseMaterialFromFileName(fileName) ?? "UNKNOWN";

                    string materialType = MaterialTypeNormalizer.NormalizeToType(materialRaw);
                    string materialTag = MaterialTypeNormalizer.MakeSafeTag(materialType);

                    string fullPath = Path.Combine(sub, fileName);
                    if (!File.Exists(fullPath))
                        continue;

                    string key = MakeKey(fileName, thicknessMm, materialTag, folderName);

                    if (!combined.TryGetValue(key, out var existing))
                    {
                        combined[key] = new CombinedPart
                        {
                            FileName = fileName,
                            FolderName = folderName,
                            FullPath = fullPath,
                            ThicknessMm = thicknessMm,
                            Quantity = Math.Max(1, qty),
                            MaterialRaw = materialRaw,
                            MaterialType = materialType,
                            MaterialTag = materialTag
                        };
                    }
                    else
                    {
                        existing.Quantity += Math.Max(1, qty);
                    }
                }
            }

            if (combined.Count == 0)
            {
                if (showUi) MessageBox.Show("No valid parts.csv entries were found.", "Combine DWG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var list = combined.Values
                .OrderBy(p => p.ThicknessMm)
                .ThenBy(p => p.MaterialTag, StringComparer.OrdinalIgnoreCase)
                .ThenBy(p => p.FileName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // all_parts.csv now includes MaterialRaw + MaterialType (extra column at end)
            string allCsvPath = Path.Combine(mainFolder, "all_parts.csv");
            var outLines = new List<string> { "FileName,PlateThickness_mm,Quantity,Folder,Material,MaterialType" };

            foreach (var p in list)
            {
                outLines.Add(string.Format(CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3},{4},{5}",
                    EscapeCsv(p.FileName),
                    p.ThicknessMm.ToString("0.###", CultureInfo.InvariantCulture),
                    p.Quantity,
                    EscapeCsv(p.FolderName),
                    EscapeCsv(p.MaterialRaw),
                    EscapeCsv(p.MaterialType)));
            }

            try { File.WriteAllLines(allCsvPath, outLines, Encoding.UTF8); } catch { }

            CreatePerThicknessDwgs(mainFolder, list);

            if (showUi)
            {
                MessageBox.Show(
                    "DWG combination finished.\r\n\r\n" +
                    "Unique part-lines: " + list.Count + "\r\n" +
                    "Summary CSV: " + allCsvPath,
                    "Combine DWG",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private static string MakeKey(string fileName, double thicknessMm, string materialTag, string folderName)
        {
            return (fileName ?? "").Trim().ToUpperInvariant() + "|" +
                   thicknessMm.ToString("0.###", CultureInfo.InvariantCulture) + "|" +
                   (materialTag ?? "OTHER").Trim().ToUpperInvariant() + "|" +
                   (folderName ?? "JOB").Trim().ToUpperInvariant();
        }

        private static string EscapeCsv(string value)
        {
            if (value == null) return "";
            bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            if (!mustQuote) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static string[] SplitCsvSimple(string line)
        {
            var cols = new List<string>();
            var sb = new StringBuilder();
            bool inQ = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQ = !inQ;
                    sb.Append(c);
                    continue;
                }

                if (c == ',' && !inQ)
                {
                    cols.Add(sb.ToString());
                    sb.Clear();
                    continue;
                }

                sb.Append(c);
            }

            cols.Add(sb.ToString());
            return cols.ToArray();
        }

        private static string ComputeFileHash8(string path)
        {
            try
            {
                using (var sha = System.Security.Cryptography.SHA256.Create())
                using (var fs = File.OpenRead(path))
                {
                    byte[] h = sha.ComputeHash(fs);
                    return BitConverter.ToString(h, 0, 4).Replace("-", ""); // 8 hex chars
                }
            }
            catch
            {
                return _rnd.Next(0, int.MaxValue).ToString("X8", CultureInfo.InvariantCulture);
            }
        }

        // Bright/distinct ACI indices (good on black background)
        private static readonly byte[] GoodAci = new byte[]
        {
            1, 2, 3, 4, 5, 6, 10, 11, 12, 13, 14, 20, 21, 30, 31, 40, 50, 60, 70, 90,
            100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250
        };

        /// <summary>
        /// Block name format used by LaserCut:
        /// P_<PartName>__MAT_<MaterialTypeTag>__H<hash8>_Q<qty>
        /// Quantity MUST remain "_Q{n}" at end.
        /// </summary>
        private static string MakePlateBlockName(string fileName, string materialTypeTag, string hash8, int quantity)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName) ?? "Part";

            // If file already has __MAT_ tag, remove it to avoid duplication
            int matIdx = baseName.LastIndexOf("__MAT_", StringComparison.OrdinalIgnoreCase);
            if (matIdx >= 0)
                baseName = baseName.Substring(0, matIdx);

            var sb = new StringBuilder();
            foreach (char c in baseName)
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');

            string safePart = sb.Length > 0 ? sb.ToString() : "Part";
            string safeMat = string.IsNullOrWhiteSpace(materialTypeTag) ? "OTHER" : materialTypeTag.Trim().ToUpperInvariant();
            string safeHash = string.IsNullOrWhiteSpace(hash8) ? "00000000" : hash8.Trim();
            int q = Math.Max(1, quantity);

            return $"P_{safePart}__MAT_{safeMat}__H{safeHash}_Q{q}";
        }

        private static string EnsureUniqueBlockName(CadDocument doc, string name)
        {
            if (doc == null || string.IsNullOrWhiteSpace(name)) return name;

            string candidate = name;
            int n = 1;

            while (doc.BlockRecords.Contains(candidate))
            {
                candidate = name + "__" + n.ToString(CultureInfo.InvariantCulture);
                n++;
            }

            return candidate;
        }

        private const double TextWidthFactor = 1.0;
        private static double EstimateTextWidth(string text, double textHeight)
        {
            if (string.IsNullOrEmpty(text) || textHeight <= 0.0) return 0.0;
            return text.Length * textHeight * TextWidthFactor;
        }

        private static void CreatePerThicknessDwgs(string mainFolder, List<CombinedPart> parts)
        {
            var thicknessGroups = parts.GroupBy(p => p.ThicknessMm).OrderBy(g => g.Key);

            foreach (var g in thicknessGroups)
            {
                double thickness = g.Key;
                string thicknessText = thickness.ToString("0.###", CultureInfo.InvariantCulture);
                string fileSafeThickness = thicknessText.Replace('.', '_').Replace(',', '_');

                string outPath = Path.Combine(mainFolder, $"thickness_{fileSafeThickness}.dwg");

                var doc = new CadDocument();
                BlockRecord modelSpace = doc.BlockRecords["*Model_Space"];

                double cursorX = 0.0;
                const double marginX = 50.0;

                int colorIndex = 0;

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

                    string hash8 = ComputeFileHash8(part.FullPath);

                    string blockName = MakePlateBlockName(part.FileName, part.MaterialTag, hash8, part.Quantity);
                    blockName = EnsureUniqueBlockName(doc, blockName);

                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    byte aci = GoodAci[colorIndex % GoodAci.Length];
                    colorIndex++;
                    var blockColor = new AcadColor(aci);

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

                    // bbox
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
                        minX = 0.0; maxX = 0.0; minY = 0.0;
                    }

                    double blockWidth = maxX - minX;
                    if (blockWidth <= 0.0) blockWidth = 1.0;

                    // text under plate
                    double textHeight = 20.0;
                    double baselineY = 0.0;
                    double gapPlateToFirst = 8.0;
                    double gapBetweenLines = 10.0;

                    double textY1 = baselineY - textHeight - gapPlateToFirst;
                    double textY2 = textY1 - textHeight - gapBetweenLines;
                    double textY3 = textY2 - textHeight - gapBetweenLines;

                    string label1 = $"Plate: {thicknessText} mm";
                    string label2 = $"Qty: {part.Quantity}";
                    string label3 = $"MatType: {part.MaterialType}";

                    double w1 = EstimateTextWidth(label1, textHeight);
                    double w2 = EstimateTextWidth(label2, textHeight);
                    double w3 = EstimateTextWidth(label3, textHeight);
                    double maxTextWidth = Math.Max(w1, Math.Max(w2, w3));

                    double extraPad = textHeight;
                    double colWidth = maxTextWidth > blockWidth ? maxTextWidth + 2 * extraPad : blockWidth;

                    double colCenterX = cursorX + colWidth / 2.0;
                    double blockCenterLocalX = (minX + maxX) * 0.5;

                    double insertX = colCenterX - blockCenterLocalX;
                    double insertY = -minY;

                    modelSpace.Entities.Add(new Insert(block)
                    {
                        InsertPoint = new XYZ(insertX, insertY, 0),
                        XScale = 1,
                        YScale = 1,
                        ZScale = 1
                    });

                    modelSpace.Entities.Add(new MText { Value = label1, InsertPoint = new XYZ(colCenterX - w1 / 2.0, textY1, 0), Height = textHeight });
                    modelSpace.Entities.Add(new MText { Value = label2, InsertPoint = new XYZ(colCenterX - w2 / 2.0, textY2, 0), Height = textHeight });
                    modelSpace.Entities.Add(new MText { Value = label3, InsertPoint = new XYZ(colCenterX - w3 / 2.0, textY3, 0), Height = textHeight });

                    cursorX += colWidth + marginX;
                }

                using (var writer = new DwgWriter(outPath, doc))
                    writer.Write();
            }
        }
    }
}
