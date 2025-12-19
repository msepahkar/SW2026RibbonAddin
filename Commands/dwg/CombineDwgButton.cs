using System;
using System.Collections.Generic;
using System.Drawing;                 // Bitmap/Graphics/Font (measurement only)
using System.Drawing.Text;
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

// DWG entity colors must be ACadSharp.Color (avoid ambiguity with System.Drawing.Color)
using AcadColor = ACadSharp.Color;

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

        // --------------------------------------------------------------------
        // 1) RANDOM COLORS that are visible on black background
        // --------------------------------------------------------------------
        private static readonly byte[][] VisibleAciGroups =
        {
            new byte[] { 11, 241 },                 // bright reds / pinkish reds
            new byte[] { 20, 30 },                  // oranges
            new byte[] { 40, 50 },                  // yellows
            new byte[] { 60, 70 },                  // yellow-green / lime
            new byte[] { 80, 90 },                  // greens
            new byte[] { 100, 110 },                // green-cyans
            new byte[] { 120, 130 },                // cyans
            new byte[] { 140, 150, 161, 171 },      // brighter blues
            new byte[] { 191, 201 },                // brighter purples
            new byte[] { 210, 220, 230 },           // magenta -> pink
            new byte[] { 7 }                        // white
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

            int groupCount = VisibleAciGroups.Length;

            int[] groupOrder = Enumerable.Range(0, groupCount).ToArray();
            Shuffle(groupOrder);

            var perGroupQueues = new Queue<byte>[groupCount];
            for (int i = 0; i < groupCount; i++)
            {
                byte[] colors = (byte[])VisibleAciGroups[i].Clone();
                Shuffle(colors);
                perGroupQueues[i] = new Queue<byte>(colors);
            }

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

        private static AcadColor NextVisibleColor()
        {
            if (_visibleAciQueue.Count == 0)
                RefillVisibleColorQueue();

            return new AcadColor(_visibleAciQueue.Dequeue());
        }

        // --------------------------------------------------------------------
        // 2) TEXT WIDTH ESTIMATION (conservative, avoids overlap in AutoCAD)
        // --------------------------------------------------------------------
        // Important: AutoCAD text width can be wider than Windows-measured Arial.
        // So we compute:
        //   width = MAX(GDI_estimate, char_count_estimate) * safety
        // This prevents under-estimation and fixes overlap.
        private const string MeasureFontFamily = "Arial";

        // Tuning knobs (safe defaults)
        private const double CharWidthFactor = 0.80;     // per-character width in "text heights"
        private const double WidthSafetyFactor = 1.20;   // extra safety to prevent overlap
        private const double SidePaddingFactor = 1.50;   // padding each side in "text heights"

        private static readonly object _measureLock = new object();
        private static Bitmap _measureBmp;
        private static Graphics _measureG;

        private static double EstimateTextWidthInDwgUnits(string text, double dwgTextHeight)
        {
            if (string.IsNullOrEmpty(text) || dwgTextHeight <= 0.0)
                return 0.0;

            // Conservative fallback: characters * height * factor
            double byChars = text.Length * dwgTextHeight * CharWidthFactor;

            double byGdi = 0.0;
            try
            {
                lock (_measureLock)
                {
                    if (_measureBmp == null)
                    {
                        _measureBmp = new Bitmap(1, 1);
                        _measureG = Graphics.FromImage(_measureBmp);
                        _measureG.TextRenderingHint = TextRenderingHint.AntiAlias;
                    }

                    // Large pixel font gives stable ratios
                    using (var font = new Font(MeasureFontFamily, 100f, FontStyle.Regular, GraphicsUnit.Pixel))
                    using (var fmt = (StringFormat)StringFormat.GenericTypographic.Clone())
                    {
                        fmt.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;

                        // Measure in pixels
                        var size = _measureG.MeasureString(text, font, int.MaxValue, fmt);

                        float h = size.Height;
                        if (h > 0.001f)
                        {
                            // Convert by ratio: width/height * desired dwg height
                            double ratio = size.Width / h;
                            byGdi = ratio * dwgTextHeight;
                        }
                    }
                }
            }
            catch
            {
                // ignore; we will fall back to byChars
                byGdi = 0.0;
            }

            double w = Math.Max(byChars, byGdi);
            return w * WidthSafetyFactor;
        }

        // --------------------------------------------------------------------
        // 3) SAFE MERGING + COLLISION RESISTANCE + LOGGING
        // --------------------------------------------------------------------
        private sealed class CombinedPart
        {
            public string FileName;
            public string FullPath;
            public double ThicknessMm;
            public int Quantity;

            // Hash used to avoid merging different DWGs that share filename/thickness
            public string HashHex;

            public HashSet<string> SourceFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public static void Combine(string mainFolder)
        {
            var log = new List<string>();

            void Log(string message)
            {
                log.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
            }

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

            // Merge key strategy:
            // - If hash exists: FileName + Thickness + Hash => merge only identical geometry
            // - If missing/unreadable: FileName + Thickness + Folder => do NOT merge across jobs
            var combined = new Dictionary<string, CombinedPart>(StringComparer.OrdinalIgnoreCase);

            int rowsRead = 0;
            int missingDwgCount = 0;

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
                catch (Exception ex)
                {
                    Log($"Failed to read CSV: '{csvPath}' => {ex.Message}");
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
                    {
                        Log($"Bad CSV row (needs 3 columns): '{csvPath}' line {i + 1}: {line}");
                        continue;
                    }

                    string fileName = (cols[0] ?? "").Trim();
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    if (!double.TryParse((cols[1] ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double tMm))
                    {
                        Log($"Bad thickness value: '{csvPath}' line {i + 1}: {line}");
                        continue;
                    }

                    if (!int.TryParse((cols[2] ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty))
                    {
                        Log($"Bad quantity value: '{csvPath}' line {i + 1}: {line}");
                        continue;
                    }

                    rowsRead++;

                    string fullPath = Path.Combine(sub, fileName);

                    string hashHex = null;
                    if (File.Exists(fullPath))
                    {
                        hashHex = TryComputeSha256Hex(fullPath);
                        if (string.IsNullOrWhiteSpace(hashHex))
                            Log($"Failed to hash DWG (will not cross-merge): '{fullPath}'");
                    }
                    else
                    {
                        missingDwgCount++;
                        Log($"Missing DWG referenced by CSV: '{fullPath}'");
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
                        // Prefer existing file if stored path is missing
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
                WriteLogIfAny(mainFolder, log);
                return;
            }

            var list = combined.Values
                .OrderBy(p => p.ThicknessMm)
                .ThenBy(p => p.FileName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Summary CSV
            string allCsvPath = Path.Combine(mainFolder, "all_parts.csv");
            var outLines = new List<string> { "FileName,PlateThickness_mm,Quantity,Folder" };

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

            try { File.WriteAllLines(allCsvPath, outLines); }
            catch (Exception ex) { Log($"Failed to write summary CSV '{allCsvPath}': {ex.Message}"); }

            // Create per-thickness DWGs
            CreatePerThicknessDwgs(mainFolder, list, Log);

            // Log if needed
            string logPath = WriteLogIfAny(mainFolder, log);

            MessageBox.Show(
                "DWG combination finished.\r\n\r\n" +
                "Rows read from CSVs: " + rowsRead + Environment.NewLine +
                "Unique parts: " + list.Count + Environment.NewLine +
                "Missing DWGs referenced by CSV: " + missingDwgCount + Environment.NewLine +
                "Summary CSV: " + allCsvPath +
                (string.IsNullOrEmpty(logPath) ? "" : (Environment.NewLine + "Log: " + logPath)),
                "Combine DWG",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string WriteLogIfAny(string mainFolder, List<string> log)
        {
            if (log == null || log.Count == 0)
                return null;

            string logPath = Path.Combine(mainFolder, "combine_log.txt");
            try
            {
                File.WriteAllLines(logPath, log);
                return logPath;
            }
            catch
            {
                return null;
            }
        }

        private static string MakeKey(string fileName, double thicknessMm, string hashHex, string folderName)
        {
            string f = (fileName ?? "").Trim().ToUpperInvariant();
            string t = thicknessMm.ToString("0.###", CultureInfo.InvariantCulture);

            if (!string.IsNullOrWhiteSpace(hashHex))
                return f + "|" + t + "|" + hashHex.Trim().ToUpperInvariant();

            return f + "|" + t + "|" + (folderName ?? "").Trim().ToUpperInvariant();
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

            string escaped = value.Replace("\"", "\"\"");
            return "\"" + escaped + "\"";
        }

        /// <summary>
        /// Block name: P_{name}[_H{hash8}]_Q{qty}
        /// Keep _Q{qty} at end so LaserCut parsing still works.
        /// </summary>
        private static string MakePlateBlockName(string fileName, string hashHex, int quantity)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Part";

            var sb = new StringBuilder();
            foreach (char c in baseName)
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');

            string safe = sb.Length > 0 ? sb.ToString() : "Part";
            int q = Math.Max(1, quantity);

            string h = null;
            if (!string.IsNullOrWhiteSpace(hashHex))
            {
                string hh = hashHex.Trim();
                h = hh.Length <= 8 ? hh : hh.Substring(0, 8);
            }

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

            // Insert suffix before _Q so qty stays parseable
            int qIdx = desiredName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            string left = qIdx >= 0 ? desiredName.Substring(0, qIdx) : desiredName;
            string right = qIdx >= 0 ? desiredName.Substring(qIdx) : "";

            for (int i = 2; i < 9999; i++)
            {
                string candidate = left + "_D" + i.ToString(CultureInfo.InvariantCulture) + right;
                if (!BlockRecordExists(doc, candidate))
                    return candidate;
            }

            return left + "_D" + Guid.NewGuid().ToString("N").Substring(0, 6) + right;
        }

        private static void CreatePerThicknessDwgs(string mainFolder, List<CombinedPart> parts, Action<string> log)
        {
            var groups = parts
                .GroupBy(p => p.ThicknessMm)
                .OrderBy(g => g.Key);

            foreach (var g in groups)
            {
                // Fresh color sequence per output DWG
                RefillVisibleColorQueue();

                double thickness = g.Key;
                string thicknessText = thickness.ToString("0.###", CultureInfo.InvariantCulture);
                string fileSafeThickness = thicknessText.Replace('.', '_').Replace(',', '_');

                string outPath = Path.Combine(mainFolder, $"thickness_{fileSafeThickness}.dwg");

                var doc = new CadDocument();

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
                    catch (Exception ex)
                    {
                        log?.Invoke($"DWG read failed: '{part.FullPath}' => {ex.Message}");
                        continue;
                    }

                    BlockRecord srcModel;
                    try
                    {
                        srcModel = srcDoc.BlockRecords["*Model_Space"];
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke($"DWG missing *Model_Space: '{part.FullPath}' => {ex.Message}");
                        continue;
                    }

                    string desiredName = MakePlateBlockName(part.FileName, part.HashHex, part.Quantity);
                    string blockName = EnsureUniqueBlockName(doc, desiredName);

                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    var blockColor = NextVisibleColor();

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

                    // labels
                    double textHeight = 20.0;

                    double baselineY = 0.0;
                    double gapPlateToFirst = 8.0;
                    double gapBetweenLines = 10.0;

                    double textY1 = baselineY - textHeight - gapPlateToFirst;
                    double textY2 = textY1 - textHeight - gapBetweenLines;

                    string label1 = $"Plate: {thicknessText} mm";
                    string label2 = $"Qty: {part.Quantity}";

                    double textWidth1 = EstimateTextWidthInDwgUnits(label1, textHeight);
                    double textWidth2 = EstimateTextWidthInDwgUnits(label2, textHeight);
                    double maxTextWidth = Math.Max(textWidth1, textWidth2);

                    // Bigger side padding to guarantee separation on black background + wide fonts
                    double extraTextSidePadding = textHeight * SidePaddingFactor;

                    double columnWidth = Math.Max(blockWidth, maxTextWidth + 2.0 * extraTextSidePadding);
                    double columnCenterX = cursorX + columnWidth / 2.0;

                    // place block centered
                    double blockCenterLocalX = (minX + maxX) * 0.5;
                    double insertX = columnCenterX - blockCenterLocalX;
                    double insertY = -minY;

                    doc.Entities.Add(new Insert(block)
                    {
                        InsertPoint = new XYZ(insertX, insertY, 0.0),
                        XScale = 1.0,
                        YScale = 1.0,
                        ZScale = 1.0
                    });

                    // center labels
                    double text1InsertX = columnCenterX - textWidth1 / 2.0;
                    double text2InsertX = columnCenterX - textWidth2 / 2.0;

                    doc.Entities.Add(new MText
                    {
                        Value = label1,
                        InsertPoint = new XYZ(text1InsertX, textY1, 0.0),
                        Height = textHeight
                    });

                    doc.Entities.Add(new MText
                    {
                        Value = label2,
                        InsertPoint = new XYZ(text2InsertX, textY2, 0.0),
                        Height = textHeight
                    });

                    cursorX += columnWidth + marginX;
                }

                try
                {
                    using (var writer = new DwgWriter(outPath, doc))
                        writer.Write();
                }
                catch (Exception ex)
                {
                    log?.Invoke($"DWG write failed: '{outPath}' => {ex.Message}");
                }
            }
        }
    }
}
