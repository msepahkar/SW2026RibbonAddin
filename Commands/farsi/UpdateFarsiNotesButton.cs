using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class UpdateFarsiNotesButton : IMehdiRibbonButton
    {
        public string Id => "UpdateFarsiNotes";

        public string DisplayName => "Update Farsi Notes";
        public string Tooltip => "Re-shape and fix all Farsi notes across all sheets";
        public string Hint => "Update Farsi Notes";

        public string SmallIconFile => "update_20.png";
        public string LargeIconFile => "update_32.png";

        public RibbonSection Section => RibbonSection.FarsiNotes;
        public int SectionOrder => 2;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing to use this command.", "Update Farsi Notes");
                    return;
                }

                var stats = context.Addin.UpdateAllFarsiNotes(model);

                MessageBox.Show(
                    $"Sheets scanned: {stats.Sheets}\r\n" +
                    $"Notes inspected: {stats.Inspected}\r\n" +
                    $"Farsi notes reshaped: {stats.Updated}\r\n" +
                    $"Skipped (non‑Farsi/no change): {stats.Skipped}",
                    "Update Farsi Notes");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while updating notes:\r\n" + ex.Message, "Update Farsi Notes");
                Debug.WriteLine("OnUpdateFarsiNotes error: " + ex);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // Only enabled for drawings
            try
            {
                var model = context.ActiveModel;
                if (model == null)
                    return AddinContext.Disable;

                return model.GetType() == (int)swDocumentTypes_e.swDocDRAWING
                    ? AddinContext.Enable
                    : AddinContext.Disable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }
    }
}
