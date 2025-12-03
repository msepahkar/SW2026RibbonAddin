using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class EditSelectedFarsiNoteButton : IMehdiRibbonButton
    {
        public string Id => "EditSelectedFarsiNote";

        public string DisplayName => "Edit Selected Note (Farsi)";
        public string Tooltip => "Open the Farsi editor for the selected note";
        public string Hint => "Edit Note (Farsi)";

        public string SmallIconFile => "edit_20.png";
        public string LargeIconFile => "edit_32.png";

        public RibbonSection Section => RibbonSection.FarsiNotes;
        public int SectionOrder => 1;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing and select a note.", "Farsi Editor");
                    return;
                }

                var selMgr = (ISelectionMgr)model.SelectionManager;
                if (selMgr == null || selMgr.GetSelectedObjectCount2(-1) < 1)
                {
                    MessageBox.Show("Select a note first.", "Farsi Editor");
                    return;
                }

                object selObj = selMgr.GetSelectedObject6(1, -1);
                INote note = null;

                if (selObj is INote n1)
                {
                    note = n1;
                }
                else if (selObj is IAnnotation ann)
                {
                    try { note = (INote)ann.GetSpecificAnnotation(); } catch { }
                }

                if (note == null)
                {
                    MessageBox.Show("The selected object is not a note.", "Farsi Editor");
                    return;
                }

                // Uses Addin's helper (changed to internal)
                context.Addin.EditNoteWithFarsiEditor(note, model);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Farsi editor: " + ex.Message, "Farsi Editor");
                Debug.WriteLine("OnEditSelectedNoteFarsi error: " + ex);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                return (model != null && model.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
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
