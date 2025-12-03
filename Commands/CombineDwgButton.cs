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
using CSMath;   // this is the math namespace (XYZ, etc.)

namespace SW2026RibbonAddin.Commands
{
    internal sealed class CombineDwgButton : IMehdiRibbonButton
    {
        public string Id => "CombineDwg";

        public string DisplayName => "Combine\nDWG";
        public string Tooltip => "Combine DWG exports from multiple jobs into per-thickness DWGs and a summary CSV.";
        public string Hint => "Combine DWG exports";

        public string SmallIconFile => "dwg_20.png";  // reuse DWG icon
        public string LargeIconFile => "dwg_32.png";

        public RibbonSection Section => RibbonSection.General;
        public int SectionOrder => 2;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolderWithFileDialog();
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
            // This command does not depend on the active SW document
            return AddinContext.Enable;
        }

        private static string SelectMainFolderWithFileDialog()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select the MAIN folder that contains job subfolders";
                dialog.Filter = "Folders|*.this_is_a_folder_selector";
                dialog.CheckFileExists = false;
                dialog.FileName = "SelectFolder";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return null;

                string path = Path.GetDirectoryName(dialog.FileName);
                return path;
            }
        }
    }

    internal static class DwgBatchCombiner
    {
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

            // ---- read all parts.csv and merge rows ----
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

            // ---- write all_parts.csv in MAIN folder ----
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

            // ---- create per-thickness DWG files ----
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

                var doc = new CadDocument(ACadVersion.AC1024);
                BlockRecord modelSpace = doc.BlockRecords["*Model_Space"];

                double offsetX = 0.0;
                const double spacingX = 1000.0;

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

                    // create a block and copy all entities into it
                    string blockName = "P_" + Guid.NewGuid().ToString("N");
                    var block = new BlockRecord(blockName);
                    doc.BlockRecords.Add(block);

                    foreach (var ent in srcModel.Entities)
                    {
                        if (ent == null)
                            continue;

                        var cloned = ent.Clone() as Entity;
                        if (cloned == null)
                            continue;

                        block.Entities.Add(cloned);
                    }

                    // insert the block
                    var insert = new Insert(block)
                    {
                        InsertPoint = new XYZ(offsetX, 0, 0),
                        XScale = 1.0,
                        YScale = 1.0,
                        ZScale = 1.0
                    };

                    doc.Entities.Add(insert);

                    // add text: Plate & Qty under the block
                    double textHeight = 20.0;
                    double gap = 5.0;

                    string label1 = $"Plate: {thicknessText} mm";
                    string label2 = $"Qty: {part.Quantity}";

                    var text1 = new MText
                    {
                        Value = label1,
                        InsertPoint = new XYZ(offsetX, -textHeight - gap, 0),
                        Height = textHeight
                    };

                    var text2 = new MText
                    {
                        Value = label2,
                        InsertPoint = new XYZ(offsetX, -2 * textHeight - 2 * gap, 0),
                        Height = textHeight
                    };

                    doc.Entities.Add(text1);
                    doc.Entities.Add(text2);

                    offsetX += spacingX;
                }

                using (var writer = new DwgWriter(outPath, doc))
                {
                    writer.Write();
                }
            }
        }
    }
}
