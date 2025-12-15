using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SW2026RibbonAddin.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class SetFastenerPropertiesButton : IMehdiRibbonButton
    {
        public string Id => "SetFastenerProperties";

        public string DisplayName => "Set fastener\nproperties...";
        public string Tooltip => "Open a dialog to set standardized fastener custom properties.";
        public string Hint => "Set fastener properties on the active part";

        public string SmallIconFile => "custom_properties_20.png";
        public string LargeIconFile => "custom_properties_32.png";

        public RibbonSection Section => RibbonSection.PartCreation;
        public int SectionOrder => 1;

        public bool IsFreeFeature => true;

        public int GetEnableState(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null)
                    return AddinContext.Disable;

                if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                    return AddinContext.Disable;

                // Only allow on *standalone* parts (not Toolbox library parts)
                var ext = model.Extension as ModelDocExtension;
                if (ext != null &&
                    ext.ToolboxPartType != (int)swToolBoxPartType_e.swNotAToolboxPart)
                {
                    return AddinContext.Disable;
                }

                return AddinContext.Enable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }

        public void Execute(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show("Open a fastener part (*.sldprt) before running this command.",
                        "Set fastener properties");
                    return;
                }

                // Ensure this is not a Toolbox library part (we only want standalone clones)
                var ext = model.Extension as ModelDocExtension;
                if (ext != null &&
                    ext.ToolboxPartType != (int)swToolBoxPartType_e.swNotAToolboxPart)
                {
                    MessageBox.Show(
                        "This part is still controlled by Toolbox.\r\n\r\n" +
                        "Use the 'Clone Part' command first to create a standalone fastener,\r\n" +
                        "then run 'Set fastener properties...'.",
                        "Set fastener properties",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // Prefill as much as possible (Toolbox-style name or our own naming)
                var values = FastenerPropertyHelper.BuildInitialValues(model);

                using (var dlg = new FastenerPropertiesForm(values))
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;
                }

                // 1) Write properties back to the model
                FastenerPropertyHelper.WriteProperties(model, values);

                // 2) Rename the file according to our naming convention
                RenameFileAccordingToConvention(model, values);

                // 3) Rebuild so any property-linked annotations update
                model.ForceRebuild3(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while setting fastener properties:\r\n\r\n" + ex.Message,
                    "Set fastener properties");
                Debug.WriteLine("SetFastenerPropertiesButton.Execute error: " + ex);
            }
        }

        private static void RenameFileAccordingToConvention(IModelDoc2 model, FastenerInitialValues values)
        {
            if (model == null || values == null)
                return;

            string currentPath;
            try
            {
                currentPath = model.GetPathName();
            }
            catch
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(currentPath))
                return;

            var folder = Path.GetDirectoryName(currentPath) ?? string.Empty;
            var ext = Path.GetExtension(currentPath);

            var newFileName = FastenerPropertyHelper.BuildStandardFileName(values, ext);
            if (string.IsNullOrWhiteSpace(newFileName))
                return;

            var targetPath = Path.Combine(folder, newFileName);

            // If already correct, nothing to do
            if (string.Equals(currentPath, targetPath, StringComparison.OrdinalIgnoreCase))
                return;

            // Do not silently overwrite another file
            if (File.Exists(targetPath))
            {
                MessageBox.Show(
                    "A file with the target name already exists:\r\n\r\n" +
                    targetPath + "\r\n\r\n" +
                    "Fastener properties were updated, but the file name was not changed.",
                    "Set fastener properties",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // SaveAs3 with Silent option effectively renames the open document to the new path.
            int saveErr = model.SaveAs3(
                targetPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

            if (saveErr != 0)
            {
                MessageBox.Show(
                    "Failed to rename the file.\r\n\r\nError code: " + saveErr,
                    "Set fastener properties",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
