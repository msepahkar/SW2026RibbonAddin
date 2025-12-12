using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Ribbon button that clones the active PART document into a new part:
    /// - Reads all items (features) from the active part's feature tree
    /// - Creates a new part document from the default part template
    /// - Copies the features into the new part one by one
    /// </summary>
    internal sealed class ClonePartButton : IMehdiRibbonButton
    {
        public string Id => "ClonePart";

        public string DisplayName => "Clone\nPart";
        public string Tooltip =>
            "Create a new part and copy all features from the active part into it.";
        public string Hint => "Clone the active part into a new document";

        // Reuse an existing icon for now; you can swap to dedicated icons later
        public string SmallIconFile => "hello_20.png";
        public string LargeIconFile => "hello_32.png";

        // Put this in the new section we added
        public RibbonSection Section => RibbonSection.PartTools;
        public int SectionOrder => 0;

        // Treat as a free feature
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
                        "Clone Part is only available when a PART document is active.",
                        "Clone Part",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                CloneActivePart(swApp, (IPartDoc)model);
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

        /// <summary>
        /// Implementation of:
        /// - read all items in the tree of the active part
        /// - create a new part
        /// - add all items one by one to the tree of the new part
        /// </summary>
        private static void CloneActivePart(SldWorks swApp, IPartDoc srcPart)
        {
            if (swApp == null || srcPart == null)
                return;

            var srcModel = (IModelDoc2)srcPart;

            // 1) Read all items (features) from the feature tree
            var features = new List<IFeature>();
            var feat = srcModel.FirstFeature();

            while (feat != null)
            {
                features.Add(feat);
                feat = feat.GetNextFeature();
            }

            if (features.Count == 0)
            {
                MessageBox.Show(
                    "The active part does not contain any features to clone.",
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // 2) Create a new part based on the user’s default part template
            string template = swApp.GetUserPreferenceStringValue(
                (int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            if (string.IsNullOrWhiteSpace(template))
            {
                MessageBox.Show(
                    "SOLIDWORKS default part template is not configured.\r\n" +
                    "Set it in Tools > Options > System Options > Default Templates.",
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var newModel = (IModelDoc2)swApp.NewDocument(
                template,
                (int)swDwgPaperSizes_e.swDwgPaperA4, // ignored for parts but required by signature
                0,
                0);

            if (newModel == null)
            {
                MessageBox.Show(
                    "Failed to create the new part document.",
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            string srcTitle = srcModel.GetTitle();
            string newTitle = newModel.GetTitle();

            try
            {
                // 3) Add all items (features) one by one to the new part
                foreach (var f in features)
                {
                    string typeName = f.GetTypeName2();

                    // Optionally skip special “folder” features such as History
                    if (ShouldSkipFeature(typeName))
                        continue;

                    srcModel.ClearSelection2(true);

                    // Select the feature in the source tree
                    if (!f.Select2(false, -1))
                        continue;

                    // Copy from source
                    srcModel.EditCopy();

                    // Paste into target
                    int actErr = 0;
                    swApp.ActivateDoc2(newTitle, false, ref actErr);
                    newModel.EditPaste();

                    // Reactivate source for the next feature
                    swApp.ActivateDoc2(srcTitle, false, ref actErr);
                }

                // Leave the cloned part active
                {
                    int actErr = 0;
                    swApp.ActivateDoc2(newTitle, false, ref actErr);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show(
                    "Error while cloning part:\r\n\r\n" + ex.Message,
                    "Clone Part",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Minimal filter to avoid copying non‑geometry “folder” features.
        /// You can expand this list as needed once you see the actual type names you get.
        /// </summary>
        private static bool ShouldSkipFeature(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return true;

            // Example minimal filtering
            switch (typeName)
            {
                case "HistoryFolder":
                    return true;

                default:
                    return false;
            }
        }
    }
}
