using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;  // ISwAddin
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SW2025RibbonAddin
{
    [ComVisible(true)]
    [Guid("9E9B0B1D-8B39-4B1F-8D77-BA0D4F1D1A21")]          // <-- keep stable
    [ProgId("SW2025RibbonAddin.Addin")]                   // <-- keep stable
    public class Addin : ISwAddin
    {
        private SldWorks _swApp;
        private int _cookie;
        private ICommandManager _cmdMgr;
        private ICommandGroup _cmdGroup;

        // Command group / tab
        private const int MAIN_CMD_GROUP_ID = 1;
        private const string MAIN_CMD_GROUP_TITLE = "Mehdi Tools";
        private const string MAIN_CMD_GROUP_TOOLTIP = "Custom tools";
        private const string TAB_NAME = "Mehdi";

        // Command indices
        private int _cmdHello;
        private int _cmdAddRtlNote;
        private int _cmdEditSelNote;
        private int _cmdRefreshRtl;

        // ========= COM REGISTRATION FOR SOLIDWORKS =========
        // These are invoked by RegAsm (or VS "Register for COM interop") to add/remove
        // the SolidWorks Add-Ins registry entries so your add-in shows in the list.

        [ComRegisterFunction]
        public static void ComRegister(Type t)
        {
            try
            {
                string clsid = "{" + t.GUID.ToString().ToUpperInvariant() + "}";

                // HKLM\SOFTWARE\SolidWorks\Addins\{GUID}
                using (var k = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\SolidWorks\Addins\" + clsid))
                {
                    if (k != null)
                    {
                        // default = 0 (DWORD) is typical; SW reads Title/Description
                        k.SetValue(null, 0, RegistryValueKind.DWord);
                        k.SetValue("Title", "SW2025 Ribbon Add-in", RegistryValueKind.String);
                        k.SetValue("Description", "Farsi RTL notes + license gate", RegistryValueKind.String);
                        // Optional: load at startup for all users
                        k.SetValue("LoadAtStartup", 0, RegistryValueKind.DWord);
                    }
                }

                // HKCU\Software\SolidWorks\AddInsStartup\{GUID} = 1 (loads for current user)
                using (var k = Registry.CurrentUser.CreateSubKey(@"Software\SolidWorks\AddInsStartup\" + clsid))
                {
                    if (k != null)
                        k.SetValue(null, 1, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("COM registration failed:\r\n" + ex.Message, "SW2025RibbonAddin");
            }
        }

        [ComUnregisterFunction]
        public static void ComUnregister(Type t)
        {
            try
            {
                string clsid = "{" + t.GUID.ToString().ToUpperInvariant() + "}";

                try { Registry.LocalMachine.DeleteSubKeyTree(@"SOFTWARE\SolidWorks\Addins\" + clsid, false); } catch { }
                try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\SolidWorks\AddInsStartup\" + clsid, false); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show("COM unregistration failed:\r\n" + ex.Message, "SW2025RibbonAddin");
            }
        }
        // ====================================================

        #region ISwAddin
        public bool ConnectToSW(object ThisSW, int cookie)
        {
            try
            {
                _swApp = (SldWorks)ThisSW;
                _cookie = cookie;

                _swApp.SetAddinCallbackInfo2(0, this, _cookie);
                _cmdMgr = _swApp.GetCommandManager(_cookie);

                // === LICENSE CHECK: block the add-in if not licensed ===
                if (!LicenseGate.EnsureLicensed(null))
                {
                    MessageBox.Show("License required or invalid. The add-in will not load.",
                        "SW2025RibbonAddin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                CreateUI();

                // Intercept built-in Edit Text on notes (so we can offer RTL editing)
                _swApp.CommandOpenPreNotify += OnCommandOpenPreNotify;

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                return false;
            }
        }

        public bool DisconnectFromSW()
        {
            try
            {
                if (_swApp != null)
                    _swApp.CommandOpenPreNotify -= OnCommandOpenPreNotify;
            }
            catch { }

            try
            {
                if (_cmdMgr != null)
                {
                    var tab = _cmdMgr.GetCommandTab((int)swDocumentTypes_e.swDocDRAWING, TAB_NAME);
                    if (tab != null) _cmdMgr.RemoveCommandTab(tab);
                    _cmdMgr.RemoveCommandGroup2(MAIN_CMD_GROUP_ID, true);
                }
            }
            catch { }

            _cmdGroup = null;
            _cmdMgr = null;
            _swApp = null;
            return true;
        }
        #endregion

        #region UI
        private void CreateUI()
        {
            int err = 0;
            bool ignorePrevious = false;

            // Some PIA builds require ref for the error parameter
            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                MAIN_CMD_GROUP_ID,
                MAIN_CMD_GROUP_TITLE,
                MAIN_CMD_GROUP_TOOLTIP,
                "", -1, ignorePrevious, ref err);

            int itemOpts = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            _cmdHello = _cmdGroup.AddCommandItem2(
                "Hello", -1, "Say Hello", "Hello",
                0, nameof(OnHello), nameof(EnableAlways), 1, itemOpts);

            _cmdAddRtlNote = _cmdGroup.AddCommandItem2(
                "Add RTL Note", -1, "Insert a right-to-left note", "Add RTL Note",
                1, nameof(OnAddRtlNote), nameof(EnableOnDrawing), 2, itemOpts);

            _cmdEditSelNote = _cmdGroup.AddCommandItem2(
                "Edit Selected Note", -1, "Edit selected note with RTL dialog", "Edit Note",
                2, nameof(OnEditSelectedNote), nameof(EnableOnSelectionIsNote), 3, itemOpts);

            _cmdRefreshRtl = _cmdGroup.AddCommandItem2(
                "Refresh RTL Notes", -1, "Re-apply RTL formatting to notes in the drawing", "Refresh Notes",
                3, nameof(OnRefreshRtlNotes), nameof(EnableOnDrawing), 4, itemOpts);

            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();

            // Add a command tab for Drawings
            var tab = _cmdMgr.GetCommandTab((int)swDocumentTypes_e.swDocDRAWING, TAB_NAME) ??
                      _cmdMgr.AddCommandTab((int)swDocumentTypes_e.swDocDRAWING, TAB_NAME);

            if (tab != null)
            {
                var box = tab.AddCommandTabBox();

                var ids = new int[]
                {
                    _cmdGroup.get_CommandID(_cmdHello),
                    _cmdGroup.get_CommandID(_cmdAddRtlNote),
                    _cmdGroup.get_CommandID(_cmdEditSelNote),
                    _cmdGroup.get_CommandID(_cmdRefreshRtl)
                };

                var textTypes = new int[]
                {
                    (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                    (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                    (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                    (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow
                };

                try { box.AddCommands(ids, textTypes); } catch { }
            }
        }
        #endregion

        #region Enables
        public int EnableAlways() => 1;

        public int EnableOnDrawing()
        {
            var md = _swApp?.IActiveDoc2 as IModelDoc2;
            return (md != null && md.GetType() == (int)swDocumentTypes_e.swDocDRAWING) ? 1 : 0;
        }

        public int EnableOnSelectionIsNote()
        {
            try
            {
                var md = _swApp?.IActiveDoc2 as IModelDoc2;
                if (md == null) return 0;
                var sm = (ISelectionMgr)md.SelectionManager;
                if (sm == null || sm.GetSelectedObjectCount2(-1) < 1) return 0;

                object obj = sm.GetSelectedObject6(1, -1);
                if (obj is INote) return 1;
                if (obj is IAnnotation an)
                {
                    try { return (an.GetSpecificAnnotation() is INote) ? 1 : 0; }
                    catch { return 0; }
                }
                return 0;
            }
            catch { return 0; }
        }
        #endregion

        #region Commands
        public int OnHello()
        {
            MessageBox.Show("Hello from Mehdi Tools!", "Hello");
            return 0;
        }

        public int OnAddRtlNote()
        {
            try
            {
                var md = _swApp?.IActiveDoc2 as IModelDoc2;
                if (md == null || md.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing first.", "Add RTL Note");
                    return 0;
                }

                string text = RtlNoteDialog.Prompt("Type the note text (RTL)", "Add RTL Note");
                if (string.IsNullOrWhiteSpace(text)) return 0;

                // Add a Right‑to‑Left mark so SolidWorks lays out correctly
                var prepared = "\u200F" + text;

                var noteObj = md.InsertNote(prepared);
                var note = noteObj as INote;
                if (note != null)
                {
                    try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }
                    md.EditRebuild3();
                }
                else
                {
                    MessageBox.Show("Could not create note.", "Add RTL Note");
                }
                return 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OnAddRtlNote: " + ex);
                return 0;
            }
        }

        public int OnEditSelectedNote()
        {
            try
            {
                var md = _swApp?.IActiveDoc2 as IModelDoc2;
                if (md == null) return 0;

                var sm = (ISelectionMgr)md.SelectionManager;
                if (sm == null || sm.GetSelectedObjectCount2(-1) < 1)
                {
                    MessageBox.Show("Select a note first.", "Edit Note");
                    return 0;
                }

                INote note = null;
                object obj = sm.GetSelectedObject6(1, -1);
                if (obj is INote n) note = n;
                else if (obj is IAnnotation an)
                {
                    try { note = (INote)an.GetSpecificAnnotation(); } catch { }
                }

                if (note == null)
                {
                    MessageBox.Show("The selection is not a note.", "Edit Note");
                    return 0;
                }

                string current = "";
                try { current = note.GetText() ?? ""; } catch { }
                if (current.Length > 0 && current[0] == '\u200F') current = current.Substring(1);

                string edited = RtlNoteDialog.Prompt("Edit the note text (RTL)", "Edit Note", current);
                if (edited == null) return 0;

                string prepared = "\u200F" + edited;
                try { note.SetText(prepared); } catch { }
                try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }
                try { md.EditRebuild3(); } catch { }

                return 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OnEditSelectedNote: " + ex);
                return 0;
            }
        }

        public int OnRefreshRtlNotes()
        {
            try
            {
                var md = _swApp?.IActiveDoc2 as IModelDoc2;
                if (md == null || md.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing first.", "Refresh RTL Notes");
                    return 0;
                }

                int updated = 0;

                var drw = (IDrawingDoc)md;
                var sheet = (ISheet)drw.GetCurrentSheet();
                var views = (object[])sheet?.GetViews() ?? Array.Empty<object>();

                foreach (IView v in views)
                {
                    var annots = (object[])v?.GetAnnotations() ?? Array.Empty<object>();
                    foreach (IAnnotation a in annots)
                    {
                        INote note = null;
                        try { note = (INote)a.GetSpecificAnnotation(); } catch { }
                        if (note == null) continue;

                        string t = "";
                        try { t = note.GetText() ?? ""; } catch { }

                        if (!ContainsRtlChar(t)) continue;

                        // Ensure RTL mark + right justification
                        if (t.Length == 0 || t[0] != '\u200F') t = "\u200F" + t;
                        try { note.SetText(t); } catch { }
                        try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }
                        updated++;
                    }
                }

                try { md.EditRebuild3(); } catch { }

                MessageBox.Show(updated == 0 ? "No RTL notes found." : $"Refreshed {updated} note(s).",
                    "Refresh RTL Notes");

                return 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OnRefreshRtlNotes: " + ex);
                return 0;
            }
        }
        #endregion

        #region Helpers
        private static bool ContainsRtlChar(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
            {
                // Arabic & Persian ranges
                if ((c >= 0x0600 && c <= 0x06FF) ||
                    (c >= 0x0750 && c <= 0x077F) ||
                    (c >= 0x08A0 && c <= 0x08FF))
                    return true;
            }
            return false;
        }

        private int OnCommandOpenPreNotify(int command, int userAction)
        {
            // swCommands_e.swCommands_Edit_Text (typical value ~1811)
            const int CMD_EDIT_TEXT = 1811;
            if (command != CMD_EDIT_TEXT) return 0;

            try
            {
                var md = _swApp?.IActiveDoc2 as IModelDoc2;
                if (md == null || md.GetType() != (int)swDocumentTypes_e.swDocDRAWING) return 0;

                var sm = (ISelectionMgr)md.SelectionManager;
                if (sm == null || sm.GetSelectedObjectCount2(-1) < 1) return 0;

                object obj = sm.GetSelectedObject6(1, -1);
                INote note = null;
                if (obj is INote n) note = n;
                else if (obj is IAnnotation an) { try { note = (INote)an.GetSpecificAnnotation(); } catch { } }
                if (note == null) return 0;

                string t = "";
                try { t = note.GetText() ?? ""; } catch { }

                if (!ContainsRtlChar(t)) return 0; // let SW do normal edit

                if (t.Length > 0 && t[0] == '\u200F') t = t.Substring(1);
                string edited = RtlNoteDialog.Prompt("Edit RTL note", "Edit Note", t);
                if (edited == null) return 1; // consume without change

                string prepared = "\u200F" + edited;
                try { note.SetText(prepared); } catch { }
                try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }
                try { md.EditRebuild3(); } catch { }

                return 1; // consume default editor
            }
            catch { return 0; }
        }
        #endregion

        #region Tiny inlined RTL dialog (no external forms)
        private sealed class RtlNoteDialog : Form
        {
            private TextBox _box;
            private Button _ok;
            private Button _cancel;

            private RtlNoteDialog(string title, string message, string initial)
            {
                Text = title;
                StartPosition = FormStartPosition.CenterParent;
                Width = 560; Height = 260;
                MinimizeBox = false; MaximizeBox = false;
                ShowInTaskbar = false;
                TopMost = true;

                var lbl = new Label
                {
                    Left = 12,
                    Top = 12,
                    Width = 520,
                    Text = message
                };
                _box = new TextBox
                {
                    Left = 12,
                    Top = 36,
                    Width = 520,
                    Height = 140,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    RightToLeft = RightToLeft.Yes,
                    Text = initial ?? ""
                };
                _ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 340, Width = 90, Top = 184 };
                _cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 442, Width = 90, Top = 184 };

                Controls.AddRange(new Control[] { lbl, _box, _ok, _cancel });
                AcceptButton = _ok; CancelButton = _cancel;
            }

            public static string Prompt(string message, string title, string initial = "")
            {
                using (var dlg = new RtlNoteDialog(title, message, initial))
                {
                    return dlg.ShowDialog() == DialogResult.OK ? dlg._box.Text : null;
                }
            }
        }
        #endregion
    }
}
