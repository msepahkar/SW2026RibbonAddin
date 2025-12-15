using System;
using System.Diagnostics;
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

        // Reuse existing screw icons
        public string SmallIconFile => "std_screw_20.png";
        public string LargeIconFile => "std_screw_32.png";

        public RibbonSection Section => RibbonSection.PartCreation;
        public int SectionOrder => 1;

        // This can be free; change to false if you want license gating
        public bool IsFreeFeature => true;

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

                // Prefill as much as possible
                var values = FastenerPropertyHelper.BuildInitialValues(model);

                using (var dlg = new FastenerPropertiesForm(values))
                {
                    var result = dlg.ShowDialog();
                    if (result != DialogResult.OK)
                        return;
                }

                // Write properties back to the model
                FastenerPropertyHelper.WriteProperties(model, values);

                // Ensure Description is not empty
                if (string.IsNullOrWhiteSpace(values.Description))
                {
                    values.Description = FastenerPropertyHelper.BuildDescription(values);
                    FastenerPropertyHelper.WriteProperties(model, values);
                }

                model.ForceRebuild3(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while setting fastener properties:\n" + ex.Message,
                    "Set fastener properties");
                Debug.WriteLine("SetFastenerPropertiesButton.Execute error: " + ex);
            }
        }
    }
}
