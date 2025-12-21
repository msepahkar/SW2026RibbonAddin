// Commands\dwg\BatchCombineNestButton.cs
// DROP-IN: replace the entire file with this one.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class BatchCombineNestButton : IMehdiRibbonButton
    {
        public string Id => "BatchCombineNest";

        public string DisplayName => "Batch\nCombine+Nest";
        public string Tooltip => "Runs Combine DWG, then batch nests all thickness_*.dwg (optionally per material).";
        public string Hint => "Combine + nest in one run";

        // Reuse existing icon(s)
        public string SmallIconFile => "combine_dwg_20.png";
        public string LargeIconFile => "combine_dwg_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 4;

        public bool IsFreeFeature => false;

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolder();
            if (string.IsNullOrWhiteSpace(mainFolder))
                return;

            if (!Directory.Exists(mainFolder))
            {
                MessageBox.Show("Folder does not exist:\r\n" + mainFolder, "Batch Combine+Nest",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 1) Combine
            try
            {
                // IMPORTANT: Combine writes all_parts.csv + thickness_*.dwg into mainFolder
                DwgBatchCombiner.Combine(mainFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Combine step failed:\r\n\r\n" + ex.Message, "Batch Combine+Nest",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2) Find thickness inputs
            var inputs = Directory.GetFiles(mainFolder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
                .Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f) ?? "";
                    return n.IndexOf("_nested", StringComparison.OrdinalIgnoreCase) < 0
                        && n.IndexOf("_nest_log", StringComparison.OrdinalIgnoreCase) < 0;
                })
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (inputs.Count == 0)
            {
                MessageBox.Show("Combine finished, but no thickness_*.dwg files were found in:\r\n" + mainFolder,
                    "Batch Combine+Nest", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3) Nest options
            LaserCutRunSettings settings;
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
            }

            // 4) Total count for progress
            int totalOverall = 0;
            foreach (var f in inputs)
            {
                try { totalOverall += Math.Max(0, DwgLaserNester.CountTotalInstances(f)); }
                catch { }
            }

            if (totalOverall <= 0)
            {
                MessageBox.Show(
                    "No plate blocks found in any thickness DWG.\r\n\r\n" +
                    "Check that Combine produced blocks starting with P_ (and ideally with __MAT_...).",
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            int placedOverall = 0;

            var summary = new StringBuilder();
            summary.AppendLine("BATCH COMBINE + NEST SUMMARY");
            summary.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            summary.AppendLine($"Folder: {mainFolder}");
            summary.AppendLine($"Thickness files: {inputs.Count}");
            summary.AppendLine();
            summary.AppendLine($"SeparateByMaterial: {settings.SeparateByMaterial}");
            summary.AppendLine($"OutputOneDwgPerMaterial: {settings.SeparateByMaterial && settings.OutputOneDwgPerMaterial}");
            summary.AppendLine($"UsePerMaterialSheetPresets: {settings.UsePerMaterialSheetPresets}");
            summary.AppendLine($"GlobalSheet: {settings.DefaultSheet}");
            summary.AppendLine();

            var errors = new List<string>();

            using (var progress = new LaserCutProgressForm(totalOverall))
            {
                progress.Show();
                Application.DoEvents();

                foreach (var input in inputs)
                {
                    string fileName = Path.GetFileName(input) ?? input;

                    try
                    {
                        progress.SetStatus("Nesting: " + fileName);

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
                Application.DoEvents();
            }

            string summaryPath = Path.Combine(mainFolder, "batch_combine_nest_summary.txt");
            try { File.WriteAllText(summaryPath, summary.ToString(), Encoding.UTF8); } catch { }

            if (errors.Count == 0)
            {
                MessageBox.Show(
                    "Batch Combine+Nest finished.\r\n\r\nSummary:\r\n" + summaryPath,
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    "Batch Combine+Nest finished with errors.\r\n\r\n" +
                    string.Join("\r\n", errors.Take(12)) +
                    (errors.Count > 12 ? "\r\n..." : "") +
                    "\r\n\r\nSummary:\r\n" + summaryPath,
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
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
}
