using System;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// UI button: prompts for a MAIN folder, then combines job DWGs into thickness_*.dwg + all_parts.csv.
    /// </summary>
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
                // showUi=true so the combiner can surface user-friendly status.
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
