using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class AddFarsiNoteButton : IMehdiRibbonButton
    {
        public string Id => "AddFarsiNote";

        public string DisplayName => "Add Farsi Note";
        public string Tooltip => "Type Persian (Farsi) text and place it as a drawing note";
        public string Hint => "Farsi Note";

        public string SmallIconFile => "farsi_20.png";
        public string LargeIconFile => "farsi_32.png";

        public RibbonSection Section => RibbonSection.FarsiNotes;
        public int SectionOrder => 0;

        // Mark as paid feature for future licensing.
        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing to use this command.", "Farsi Note");
                    return;
                }

                // New-note mode: full formatting (font, size, alignment)
                using (var dlg = new Forms.FarsiNoteForm(true))
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    string text = dlg.NoteText ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(text)) return;

                    text = ArabicTextUtils.PrepareForSolidWorks(text, dlg.UseRtlMarkers, dlg.InsertJoiners);

                    // Uses Addin's helper (changed to internal)
                    context.Addin.StartFarsiNotePlacement(
                        model,
                        text,
                        dlg.SelectedFontName,
                        dlg.FontSizePoints,
                        dlg.Alignment);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show(
                    "Unexpected error while preparing note placement.\r\n" + ex.Message,
                    "Farsi Note");
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // Only enabled for drawings (better UX: no pointless clicking in Part/Assembly)
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
