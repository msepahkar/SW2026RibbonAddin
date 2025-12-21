using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nnesting";
        public string Tooltip => "Nest thickness_*.dwg into sheets (per material / per thickness).";
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

            // 1) Scan available (Material x Thickness) jobs BEFORE showing options
            List<LaserNestJob> jobs;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                jobs = DwgLaserNester.ScanJobsForFolder(folder);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            if (jobs == null || jobs.Count == 0)
            {
                MessageBox.Show(
                    "No parts found in thickness_*.dwg files.\r\n\r\nMake sure you ran Combine DWG first.",
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // default run settings (the 3 checkboxes are ALWAYS enabled now)
            LaserCutRunSettings settings;
            List<LaserNestJob> selectedJobs;

            using (var dlg = new LaserCutOptionsForm(folder, jobs))
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                settings = dlg.Settings;
                selectedJobs = dlg.SelectedJobs;
            }

            if (selectedJobs == null || selectedJobs.Count == 0)
            {
                MessageBox.Show(
                    "Nothing selected.\r\n\r\nSelect at least one thickness/material to run nesting.",
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                DwgLaserNester.NestJobs(folder, selectedJobs, settings, showUi: true);
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
