using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class BatchCombineNestButton : IMehdiRibbonButton
    {
        public string Id => "BatchCombineNest";

        public string DisplayName => "Batch\nCombine+Nest";
        public string Tooltip => "Runs Combine DWG (builds thickness_*.dwg) then runs Laser nesting on the same folder.";
        public string Hint => "Combine then Nest";

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
                MessageBox.Show(
                    "Folder does not exist:\r\n" + mainFolder,
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 1) Combine (creates all_parts.csv + thickness_*.dwg in mainFolder)
                var combineResult = DwgBatchCombiner.Combine(mainFolder, showUi: false);

                // 2) Nest options
                LaserCutRunSettings settings;
                using (var dlg = new LaserCutOptionsForm())
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;

                    settings = dlg.Settings;
                }

                // 3) Nest folder (reads thickness_*.dwg and outputs nested DWGs)
                DwgLaserNester.NestFolder(mainFolder, settings, showUi: false);

                // Final message (single message for the whole pipeline)
                var sb = new StringBuilder();
                sb.AppendLine("Batch Combine+Nest finished.");
                sb.AppendLine();
                sb.AppendLine("Folder:");
                sb.AppendLine(mainFolder);
                sb.AppendLine();
                sb.AppendLine("Outputs:");
                sb.AppendLine("- all_parts.csv");
                sb.AppendLine("- thickness_*.dwg");
                sb.AppendLine("- thickness_*_nested_*.dwg");
                sb.AppendLine("- batch_nest_summary.txt");

                MessageBox.Show(
                    sb.ToString(),
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Batch Combine+Nest failed:\r\n\r\n" + ex.Message,
                    "Batch Combine+Nest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
