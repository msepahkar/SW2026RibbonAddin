using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Step 7: One-click workflow after exports exist:
    /// - Pick MAIN folder
    /// - Combine DWGs (calls DwgBatchCombiner.Combine)
    /// - Nest all thickness_*.dwg (silent) using ONE options form
    /// - Write ONE summary file + show ONE message box
    /// </summary>
    internal sealed class BatchCombineNestButton : IMehdiRibbonButton
    {
        public string Id => "BatchCombineNest";

        public string DisplayName => "Batch\nNest";
        public string Tooltip => "Combine + nest all thickness DWGs in a main folder (one run).";
        public string Hint => "Combine and nest all";

        // Reuse existing icons to avoid adding resources
        public string SmallIconFile => "combine_dwg_20.png";
        public string LargeIconFile => "combine_dwg_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 4;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolder();
            if (string.IsNullOrEmpty(mainFolder))
                return;

            // Let user choose whether to run Combine first (keeps your “opinion” in the workflow)
            var choice = MessageBox.Show(
                "Run Combine DWG first?\r\n\r\n" +
                "Yes  = Combine + Nest all thickness files\r\n" +
                "No   = Only Nest existing thickness files\r\n" +
                "Cancel = Stop",
                "Batch Nest",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (choice == DialogResult.Cancel)
                return;

            if (choice == DialogResult.Yes)
            {
                try
                {
                    DwgBatchCombiner.Combine(mainFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Combine failed:\r\n\r\n" + ex.Message, "Batch Nest",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // Find thickness DWGs (skip already-nested outputs)
            var thicknessFiles = Directory.GetFiles(mainFolder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
                .Where(p =>
                {
                    string n = Path.GetFileNameWithoutExtension(p) ?? "";
                    return n.IndexOf("_nested", StringComparison.OrdinalIgnoreCase) < 0;
                })
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (thicknessFiles.Count == 0)
            {
                MessageBox.Show("No thickness_*.dwg files were found in the selected folder.", "Batch Nest",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            double sheetWidth;
            double sheetHeight;
            RotationMode rotationMode;
            int anyAngleStepDeg;
            bool writeReportCsv;

            using (var dlg = new LaserCutOptionsForm())
            {
                dlg.Text = "Batch nest options (applies to ALL thickness files)";
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                sheetWidth = dlg.SheetWidthMm;
                sheetHeight = dlg.SheetHeightMm;
                rotationMode = dlg.RotationMode;
                anyAngleStepDeg = dlg.AnyAngleStepDeg;
                writeReportCsv = dlg.WriteReportCsv;
            }

            var results = new List<DwgLaserNester.NestResult>();
            var errors = new List<string>();

            using (var progress = new SimpleBatchProgressForm(thicknessFiles.Count))
            {
                progress.Show();
                Application.DoEvents();

                for (int i = 0; i < thicknessFiles.Count; i++)
                {
                    string f = thicknessFiles[i];
                    progress.Step($"Nesting {i + 1} / {thicknessFiles.Count}\r\n{Path.GetFileName(f)}");

                    try
                    {
                        // Silent mode: no per-file MessageBox
                        var r = DwgLaserNester.Nest(
                            f, sheetWidth, sheetHeight,
                            rotationMode, anyAngleStepDeg,
                            writeReportCsv,
                            showUi: false);

                        results.Add(r);
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{Path.GetFileName(f)} => {ex.Message}");
                    }
                }

                progress.Close();
            }

            // Write summary file
            string summaryPath = Path.Combine(mainFolder, "batch_nest_summary.txt");
            TryWriteSummary(summaryPath, mainFolder, thicknessFiles, results, errors);

            // One final message box
            MessageBox.Show(
                "Batch nest finished.\r\n\r\n" +
                $"Thickness files found: {thicknessFiles.Count}\r\n" +
                $"Nested successfully: {results.Count}\r\n" +
                $"Failed: {errors.Count}\r\n\r\n" +
                $"Summary: {summaryPath}",
                "Batch Nest",
                MessageBoxButtons.OK,
                errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        private static string SelectMainFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the MAIN folder that contains job subfolders (parts.csv + DWGs)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }

        private static void TryWriteSummary(
            string summaryPath,
            string mainFolder,
            List<string> thicknessFiles,
            List<DwgLaserNester.NestResult> results,
            List<string> errors)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Batch Nest Summary");
                sb.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Main folder: {mainFolder}");
                sb.AppendLine();

                sb.AppendLine("=== Results ===");
                foreach (var r in results.OrderBy(x => x.SourceDwgPath, StringComparer.OrdinalIgnoreCase))
                {
                    sb.AppendLine($"Source: {Path.GetFileName(r.SourceDwgPath)}");
                    sb.AppendLine($"  Sheets: {r.SheetsUsed}");
                    sb.AppendLine($"  Parts: {r.TotalParts}");
                    sb.AppendLine($"  Blocks found/skipped: {r.CandidateBlocks}/{r.SkippedBlocks}");
                    sb.AppendLine($"  Output: {r.OutputDwgPath}");
                    if (!string.IsNullOrEmpty(r.ReportCsvPath)) sb.AppendLine($"  Report: {r.ReportCsvPath}");
                    if (!string.IsNullOrEmpty(r.LogPath)) sb.AppendLine($"  Log: {r.LogPath}");
                    sb.AppendLine();
                }

                if (errors.Count > 0)
                {
                    sb.AppendLine("=== Errors ===");
                    foreach (var e in errors)
                        sb.AppendLine(e);
                    sb.AppendLine();
                }

                sb.AppendLine("=== Thickness DWG files scanned ===");
                foreach (var f in thicknessFiles)
                    sb.AppendLine(Path.GetFileName(f));

                File.WriteAllText(summaryPath, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // ignore
            }
        }

        private sealed class SimpleBatchProgressForm : Form
        {
            private readonly ProgressBar _bar;
            private readonly Label _label;
            private readonly int _max;
            private int _cur;

            public SimpleBatchProgressForm(int maximum)
            {
                if (maximum <= 0) maximum = 1;
                _max = maximum;

                Text = "Batch nest...";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MinimizeBox = false;
                MaximizeBox = false;
                ShowInTaskbar = false;
                ClientSize = new System.Drawing.Size(420, 110);

                _label = new Label
                {
                    AutoSize = false,
                    Location = new System.Drawing.Point(12, 10),
                    Size = new System.Drawing.Size(396, 40),
                    Text = "Starting..."
                };
                Controls.Add(_label);

                _bar = new ProgressBar
                {
                    Location = new System.Drawing.Point(12, 60),
                    Size = new System.Drawing.Size(396, 20),
                    Minimum = 0,
                    Maximum = _max,
                    Value = 0
                };
                Controls.Add(_bar);
            }

            public void Step(string text)
            {
                _label.Text = text ?? "";
                if (_cur < _max)
                {
                    _cur++;
                    _bar.Value = _cur;
                }
                _bar.Refresh();
                _label.Refresh();
                Application.DoEvents();
            }
        }
    }
}
