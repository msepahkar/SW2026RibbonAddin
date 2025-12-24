using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class BatchCombineNestButton : IMehdiRibbonButton
    {
        public string Id => "BatchCombineNest";

        public string DisplayName => "Batch\nCombine+Nest";
        public string Tooltip => "Runs Combine DWG on a main folder, then nests thickness_*.dwg per exact material + thickness.";
        public string Hint => "Batch combine + nest";

        // Reuse combine icons (change if you want)
        public string SmallIconFile => "combine_dwg_20.png";
        public string LargeIconFile => "combine_dwg_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 4;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string mainFolder = SelectMainFolder();
            if (string.IsNullOrWhiteSpace(mainFolder))
                return;

            try
            {
                // 1) Combine (with progress)
                DwgBatchCombiner.CombineRunResult combine;

                using (var prog = new DwgCombineProgressForm())
                {
                    prog.Show();
                    combine = DwgBatchCombiner.Combine(mainFolder, showUi: false, progress: prog);
                    try { prog.Close(); } catch { }
                }
                if (combine != null && !string.IsNullOrWhiteSpace(combine.ErrorMessage))
                {
                    MessageBox.Show(
                        combine.ErrorMessage,
                        "Batch Combine + Nest",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // all_parts.csv is regenerated during Combine. Clear the index cache so the subsequent
                // ScanJobsForFolder call always picks up the freshly written CSV (important on file
                // systems with coarse timestamp resolution).
                DwgLaserNester.ClearAllPartsIndexCacheForFolder(mainFolder);

                // 2) Scan nesting jobs (Material x Thickness)
                var jobs = DwgLaserNester.ScanJobsForFolder(mainFolder);
                if (jobs == null || jobs.Count == 0)
                {
                    MessageBox.Show(
                        "No thickness_*.dwg files were found after combining.\r\n\r\n" +
                        "Make sure the main folder contains job subfolders with parts.csv + DWGs.",
                        "Batch Combine + Nest",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // 3) Show the new options dialog (needs folder + jobs)
                LaserCutRunSettings settings;
                List<LaserNestJob> selectedJobs;

                using (var dlg = new LaserCutOptionsForm(mainFolder, jobs))
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;

                    settings = dlg.Settings;
                    selectedJobs = dlg.SelectedJobs;
                }

                if (selectedJobs == null || selectedJobs.Count == 0)
                {
                    MessageBox.Show(
                        "Nothing selected.",
                        "Batch Combine + Nest",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // 4) Run nesting batch
                DwgLaserNester.NestJobs(mainFolder, selectedJobs, settings, showUi: true);
            }
            catch (OperationCanceledException)
            {
                // User cancelled DWG combining.
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Batch Combine + Nest failed:\r\n\r\n" + ex.Message,
                    "Batch Combine + Nest",
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
}
