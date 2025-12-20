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

        public string DisplayName => "Batch\nNest";
        public string Tooltip => "Combine + nest all thickness_*.dwg in the selected main folder.";
        public string Hint => "Batch combine + nest";

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

            var choice = MessageBox.Show(
                "Run Combine DWG first?\r\n\r\n" +
                "Yes  = Combine + Nest\r\n" +
                "No   = Nest only existing thickness_*.dwg\r\n" +
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
                    DwgBatchCombiner.Combine(mainFolder, showUi: false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Combine DWG failed:\r\n\r\n" + ex.Message, "Batch Nest",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

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
                MessageBox.Show("No thickness_*.dwg files found in this folder.", "Batch Nest",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            LaserCutRunSettings settings;
            using (var dlg = new LaserCutOptionsForm())
            {
                dlg.Text = "Batch nest options (applies to ALL thickness files)";
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
            }

            var results = new List<(string thicknessFile, DwgLaserNester.NestRunResult result)>();
            var errors = new List<string>();

            using (var prog = new SimpleBatchProgressForm(thicknessFiles.Count))
            {
                prog.Show();
                Application.DoEvents();

                for (int i = 0; i < thicknessFiles.Count; i++)
                {
                    string f = thicknessFiles[i];
                    prog.Step($"Nesting {i + 1}/{thicknessFiles.Count}\r\n{Path.GetFileName(f)}");

                    try
                    {
                        var r = DwgLaserNester.Nest(f, settings, showUi: false);
                        results.Add((f, r));
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{Path.GetFileName(f)} => {ex.Message}");
                    }
                }

                prog.Close();
            }

            string summaryPath = Path.Combine(mainFolder, "batch_nest_summary.txt");
            TryWriteSummary(summaryPath, mainFolder, thicknessFiles, results, errors);

            MessageBox.Show(
                "Batch nest finished.\r\n\r\n" +
                $"Thickness files: {thicknessFiles.Count}\r\n" +
                $"OK: {results.Count}\r\n" +
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
                dlg.Description = "Select MAIN folder (contains job subfolders with parts.csv + DWGs)";
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
            List<(string thicknessFile, DwgLaserNester.NestRunResult result)> results,
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
                foreach (var item in results.OrderBy(x => x.thicknessFile, StringComparer.OrdinalIgnoreCase))
                {
                    var r = item.result;

                    sb.AppendLine($"Thickness DWG: {Path.GetFileName(item.thicknessFile)}");
                    sb.AppendLine($"  SeparateByMaterial: {r.Settings.SeparateByMaterial}");
                    sb.AppendLine($"  OneDwgPerMaterial: {r.Settings.OutputOneDwgPerMaterial}");
                    sb.AppendLine($"  PerMaterialSheets: {r.Settings.UsePerMaterialSheetPresets}");
                    sb.AppendLine($"  Blocks found/skipped: {r.CandidateBlocks}/{r.SkippedBlocks}");

                    foreach (var o in r.Outputs)
                    {
                        sb.AppendLine($"    [{o.MaterialType}] Sheets: {o.SheetsUsed} Parts: {o.TotalParts} Sheet: {o.Sheet}");
                        sb.AppendLine($"       {o.OutputDwgPath}");
                    }

                    if (!string.IsNullOrEmpty(r.LogPath))
                        sb.AppendLine($"  Log: {r.LogPath}");

                    sb.AppendLine();
                }

                if (errors.Count > 0)
                {
                    sb.AppendLine("=== Errors ===");
                    foreach (var e in errors) sb.AppendLine(e);
                    sb.AppendLine();
                }

                sb.AppendLine("=== Thickness DWGs scanned ===");
                foreach (var f in thicknessFiles) sb.AppendLine(Path.GetFileName(f));

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

                Text = "Batch nesting...";
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
