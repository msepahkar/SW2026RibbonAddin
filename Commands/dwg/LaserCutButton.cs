using System;
using System.IO;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nnesting";
        public string Tooltip => "Nest thickness_*.dwg files into sheets (Fast / Contour L1 / Contour L2).";
        public string Hint => "Laser cut nesting";

        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string folder = SelectFolder();
            if (string.IsNullOrWhiteSpace(folder))
                return;

            LaserCutRunSettings settings;
            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
            }

            try
            {
                DwgLaserNester.NestFolder(folder, settings, showUi: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Laser nesting failed:\r\n\r\n" + ex.Message,
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;

        private static string SelectFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder containing thickness_*.dwg files (output of Combine DWG)";
                dlg.ShowNewFolderButton = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                if (!Directory.Exists(dlg.SelectedPath))
                    return null;

                return dlg.SelectedPath;
            }
        }
    }
}
