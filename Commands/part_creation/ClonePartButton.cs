using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class ClonePartButton : IMehdiRibbonButton
    {
        public string Id => "ClonePart";
        public string DisplayName => "Clone\nPart";
        public string Tooltip => "Create a new copy of the active part (same feature tree, new file).";
        public string Hint => "Clone the active part into a new part file";

        public string SmallIconFile => "clone_20.png";
        public string LargeIconFile => "clone_32.png";

        public RibbonSection Section => RibbonSection.PartCreation;
        public int SectionOrder => 0;
        public bool IsFreeFeature => true;

        public void Execute(AddinContext context)
        {
            try
            {
                var swApp = context.SwApp;
                var model = context.ActiveModel;

                if (swApp == null || model == null)
                {
                    MessageBox.Show(
                        "No active document. Open a part first.",
                        "Clone Part",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show(
                        "Clone Part is only available for part documents.",
                        "Clone Part",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                ClonePartUsingSaveAsCopy(swApp, (IPartDoc)model);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show(
                    "Unexpected error while cloning part:\r\n\r\n" + ex.Message,
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null)
                    return AddinContext.Disable;

                return model.GetType() == (int)swDocumentTypes_e.swDocPART
                    ? AddinContext.Enable
                    : AddinContext.Disable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }

        private static void ClonePartUsingSaveAsCopy(SldWorks swApp, IPartDoc srcPart)
        {
            if (swApp == null || srcPart == null)
                return;

            var srcModel = (IModelDoc2)srcPart;
            string srcPath = srcModel.GetPathName();

            if (string.IsNullOrWhiteSpace(srcPath))
            {
                MessageBox.Show(
                    "Please save the part once before cloning it.",
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string folder = Path.GetDirectoryName(srcPath) ?? "";
            string baseName = Path.GetFileNameWithoutExtension(srcPath) ?? "Part";
            string ext = Path.GetExtension(srcPath);
            string defaultCloneName = baseName + "_Clone" + ext;

            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Clone Part - choose file name";
                dlg.InitialDirectory = folder;
                dlg.Filter = "SOLIDWORKS Part (*.sldprt)|*.sldprt|All files (*.*)|*.*";
                dlg.FileName = defaultCloneName;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                string clonePath = dlg.FileName;

                // 1) Save a copy of the current part
                int saveErr = srcModel.SaveAs3(
                    clonePath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Copy);

                // SaveAs3 returns 0 on success
                if (saveErr != 0)
                {
                    MessageBox.Show(
                        "Failed to create cloned part.\r\n\r\nError code: " + saveErr,
                        "Clone Part",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 2) Open the cloned part
                int openErrors = 0;
                int openWarnings = 0;

                var newModel = (IModelDoc2)swApp.OpenDoc6(
                    clonePath,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref openErrors,
                    ref openWarnings);

                if (newModel == null || openErrors != 0)
                {
                    MessageBox.Show(
                        "Clone was saved but could not be opened.\r\n\r\nError code: " + openErrors,
                        "Clone Part",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 3) Force the clone to be a NON‑Toolbox part and save once
                try
                {
                    var newExt = (ModelDocExtension)newModel.Extension;

                    // swNotAToolboxPart (0) = Not a Toolbox part
                    newExt.ToolboxPartType = (int)swToolBoxPartType_e.swNotAToolboxPart;

                    int saveErrors2 = 0;
                    int saveWarnings2 = 0;
                    bool saveOk = newModel.Save3(
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref saveErrors2,
                        ref saveWarnings2);

                    if (!saveOk || saveErrors2 != 0)
                    {
                        MessageBox.Show(
                            "Clone was created but could not be fully de‑linked from Toolbox.\r\n\r\nSave error code: " +
                            saveErrors2,
                            "Clone Part",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    // If ToolboxPartType is not available for some reason, we just log and continue
                    Debug.WriteLine(ex);
                }

                // 4) Activate the cloned part
                int actErr = 0;
                swApp.ActivateDoc2(newModel.GetTitle(), false, ref actErr);
            }
        }
    }
}
