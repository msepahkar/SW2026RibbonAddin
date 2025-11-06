using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;  // ISwAddin
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Avoid collisions with System.Windows.Forms types
using SwMouse = SolidWorks.Interop.sldworks.Mouse;
using SwCommandTab = SolidWorks.Interop.sldworks.CommandTab;
using SwCommandTabBox = SolidWorks.Interop.sldworks.CommandTabBox;

// ---- Licensing ----
using SW2025RibbonAddin.Licensing;

namespace SW2025RibbonAddin
{
    [ComVisible(true)]
    [Guid("B67E2D5A-8C73-4A3E-93B6-1761C1A8C0C5")]
    [ProgId("SW2025RibbonAddin.Addin")]
    public class Addin : ISwAddin
    {
        private const string AddinTitle = "SW2025RibbonAddin";

        private SldWorks _swApp;
        private int _cookie;
        private ICommandManager _cmdMgr;
        private ICommandGroup _cmdGroup;

        private string _smallIconPath;
        private string _largeIconPath;

        // Command indices
        private int _helloCmdIndex = -1;
        private int _farsiCmdIndex = -1;
        private int _editSelNoteCmdIndex = -1;
        private int _updateAllNotesCmdIndex = -1;
        private int _registerCmdIndex = -1;   // NEW

        // Command group / UI constants
        private const int MAIN_CMD_GROUP_ID = 1;
        private const string MAIN_CMD_GROUP_TITLE = "Mehdi Tools";
        private const string MAIN_CMD_GROUP_TOOLTIP = "Custom tools";
        private const string TAB_NAME = "Mehdi";

        private const string HELLO_CMD_NAME = "Hello";
        private const string HELLO_CMD_TOOLTIP = "Show a hello message";
        private const string HELLO_CMD_HINT = "Hello";

        private const string FARSI_CMD_NAME = "Add Farsi Note";
        private const string FARSI_CMD_TOOLTIP = "Open a dialog to type Persian (Farsi) text and place it as a drawing note";
        private const string FARSI_CMD_HINT = "Farsi Note";

        private const string EDIT_SEL_NOTE_CMD_NAME = "Edit Selected Note (Farsi)";
        private const string EDIT_SEL_NOTE_CMD_TOOLTIP = "Open the Farsi editor for the selected note";
        private const string EDIT_SEL_NOTE_CMD_HINT = "Edit Note (Farsi)";

        private const string UPDATE_ALL_NOTES_CMD_NAME = "Update Farsi Notes";
        private const string UPDATE_ALL_NOTES_CMD_TOOLTIP = "Re-shape and fix all Farsi notes across all sheets";
        private const string UPDATE_ALL_NOTES_CMD_HINT = "Update Farsi Notes";

        private const string REGISTER_CMD_NAME = "Register";                 // NEW
        private const string REGISTER_CMD_TOOLTIP = "Activate this add-in";  // NEW
        private const string REGISTER_CMD_HINT = "Register";                 // NEW

        // Tag our notes so the editor hotkey can detect them
        internal const string FARSI_NOTE_TAG_PREFIX = "MEHDI_FARSI_NOTE";

        // Enable/disable helpers
        private const int SW_ENABLE = 1;
        private const int SW_DISABLE = 0;

        // SW command we intercept (Edit Text)
        private const int CMD_EDIT_TEXT = 1811; // swCommands_e.swCommands_Edit_Text

        // Active placement session
        private FarsiNotePlacementSession _activePlacement;

        #region ISwAddin
        public bool ConnectToSW(object ThisSW, int cookie)
        {
            try
            {
                _swApp = (SldWorks)ThisSW;
                _cookie = cookie;

                _swApp.SetAddinCallbackInfo2(0, this, _cookie);
                _cmdMgr = _swApp.GetCommandManager(_cookie);

                CreateUI();

                // Intercept default "Edit Text" on notes we created
                _swApp.CommandOpenPreNotify += OnCommandOpenPreNotify;

                // One-time tip if not activated (non-destructive)
                VerifiedLicense lic; string why;
                if (!LicenseGate.IsActivated(out lic, out why))
                {
                    _swApp.SendMsgToUser2(
                        "Not activated. Use Mehdi → Register to activate.",
                        (int)swMessageBoxIcon_e.swMbInformation,
                        (int)swMessageBoxBtn_e.swMbOk);
                }

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
            catch { /* ignore */ }

            try
            {
                if (_cmdMgr != null && _cmdGroup != null)
                    _cmdMgr.RemoveCommandGroup(MAIN_CMD_GROUP_ID);
            }
            catch { /* ignore */ }

            _cmdGroup = null;
            _cmdMgr = null;
            _swApp = null;

            return true;
        }
        #endregion

        #region UI
        private void CreateUI()
        {
            int errors = 0;
            const bool ignorePrevious = true;

            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                MAIN_CMD_GROUP_ID, MAIN_CMD_GROUP_TITLE,
                MAIN_CMD_GROUP_TOOLTIP, "", -1,
                ignorePrevious, ref errors);

            // toolbar/menu icons (yours)
            _smallIconPath = ExtractResourceToFile("SW2025RibbonAddin.Resources.icon_20.png");
            _largeIconPath = ExtractResourceToFile("SW2025RibbonAddin.Resources.icon_32.png");
            try
            {
                _cmdGroup.IconList = _smallIconPath;
                _cmdGroup.MainIconList = _largeIconPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not set icon lists: {ex.Message}");
            }

            int itemOpts = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            // 1) Hello
            _helloCmdIndex = _cmdGroup.AddCommandItem2(
                HELLO_CMD_NAME, -1, HELLO_CMD_TOOLTIP, HELLO_CMD_HINT,
                0, nameof(OnHello), nameof(OnHelloEnable), 1, itemOpts);

            // 2) Add Farsi Note
            _farsiCmdIndex = _cmdGroup.AddCommandItem2(
                FARSI_CMD_NAME, -1, FARSI_CMD_TOOLTIP, FARSI_CMD_HINT,
                1, nameof(OnAddFarsiNote), nameof(OnAddFarsiNoteEnable), 2, itemOpts);

            // 3) Edit selected note (Farsi)
            _editSelNoteCmdIndex = _cmdGroup.AddCommandItem2(
                EDIT_SEL_NOTE_CMD_NAME, -1, EDIT_SEL_NOTE_CMD_TOOLTIP, EDIT_SEL_NOTE_CMD_HINT,
                2, nameof(OnEditSelectedNoteFarsi), nameof(OnEditSelectedNoteFarsiEnable), 3, itemOpts);

            // 4) Update all Farsi notes
            _updateAllNotesCmdIndex = _cmdGroup.AddCommandItem2(
                UPDATE_ALL_NOTES_CMD_NAME, -1, UPDATE_ALL_NOTES_CMD_TOOLTIP, UPDATE_ALL_NOTES_CMD_HINT,
                3, nameof(OnUpdateFarsiNotes), nameof(OnUpdateFarsiNotesEnable), 4, itemOpts);

            // 5) Register (NEW)
            _registerCmdIndex = _cmdGroup.AddCommandItem2(
                REGISTER_CMD_NAME, -1, REGISTER_CMD_TOOLTIP, REGISTER_CMD_HINT,
                4, nameof(OnRegister), nameof(OnRegisterEnable), 5, itemOpts);

            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();

            // Command Tab (DRAWING) – same as your file, plus Register button
            try
            {
                int docType = (int)swDocumentTypes_e.swDocDRAWING;

                ICommandTab tab = _cmdMgr.GetCommandTab(docType, TAB_NAME);
                if (tab != null)
                {
                    try { _cmdMgr.RemoveCommandTab((SwCommandTab)tab); } catch { /* ignore */ }
                    tab = null;
                }

                tab = _cmdMgr.AddCommandTab(docType, TAB_NAME);
                if (tab != null)
                {
                    SwCommandTabBox box = tab.AddCommandTabBox();

                    var ids = new int[]
                    {
                        _cmdGroup.get_CommandID(_helloCmdIndex),
                        _cmdGroup.get_CommandID(_farsiCmdIndex),
                        _cmdGroup.get_CommandID(_editSelNoteCmdIndex),
                        _cmdGroup.get_CommandID(_updateAllNotesCmdIndex),
                        _cmdGroup.get_CommandID(_registerCmdIndex)    // NEW
                    };

                    var textTypes = new int[]
                    {
                        (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                        (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                        (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                        (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                        (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow
                    };

                    box.AddCommands(ids, textTypes);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to create tab: " + ex.Message);
            }
        }

        private string ExtractResourceToFile(string resourceName)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var s = asm.GetManifestResourceStream(resourceName))
                {
                    if (s == null) return null;
                    var tmp = Path.Combine(Path.GetTempPath(), "SW2025_" + Path.GetFileName(resourceName));
                    using (var f = File.Create(tmp)) { s.CopyTo(f); }
                    return tmp;
                }
            }
            catch { return null; }
        }
        #endregion

        #region Commands
        public void OnHello()
        {
            try { MessageBox.Show("Hello from Mehdi Tools ✨", AddinTitle); }
            catch (Exception ex) { Debug.WriteLine(ex.ToString()); }
        }
        public int OnHelloEnable() => SW_ENABLE;

        /// <summary>Gate that asks the user to activate when needed (no destructive changes).</summary>
        private bool RequireLicense()
        {
            VerifiedLicense lic; string why;
            if (LicenseGate.IsActivated(out lic, out why)) return true;

            var r = MessageBox.Show(
                "This feature requires activation.\r\nOpen the registration dialog now?",
                AddinTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (r == DialogResult.Yes)
                LicensingUI.ShowRegistrationDialog(null);

            return false;
        }

        /// <summary>
        /// Show the Farsi editor, then start an interactive placement session.
        /// Left-click places the note at the clicked sheet coords; right-click cancels.
        /// </summary>
        public void OnAddFarsiNote()
        {
            if (!RequireLicense()) return;  // NEW

            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing to use this command.", "Farsi Note");
                    return;
                }

                using (var dlg = new Forms.FarsiNoteForm())
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    string text = dlg.NoteText ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(text)) return;

                    text = ArabicTextUtils.PrepareForSolidWorks(text, dlg.UseRtlMarkers, dlg.InsertJoiners);
                    StartFarsiNotePlacement(model, text, dlg.SelectedFontName, dlg.FontSizePoints);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("Unexpected error while preparing note placement.\r\n" + ex.Message, "Farsi Note");
            }
        }

        public int OnAddFarsiNoteEnable()
        {
            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                if (model == null) return SW_DISABLE;
                return model.GetType() == (int)swDocumentTypes_e.swDocDRAWING ? SW_ENABLE : SW_DISABLE;
            }
            catch { return SW_DISABLE; }
        }

        // Edit the currently selected note with our Farsi editor
        public void OnEditSelectedNoteFarsi()
        {
            if (!RequireLicense()) return;  // NEW

            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
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

                if (selObj is INote n1) note = n1;
                else if (selObj is IAnnotation ann)
                {
                    try { note = (INote)ann.GetSpecificAnnotation(); } catch { }
                }

                if (note == null)
                {
                    MessageBox.Show("The selected object is not a note.", "Farsi Editor");
                    return;
                }

                EditNoteWithFarsiEditor(note, model);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Farsi editor: " + ex.Message, "Farsi Editor");
                Debug.WriteLine("OnEditSelectedNoteFarsi error: " + ex);
            }
        }

        public int OnEditSelectedNoteFarsiEnable()
        {
            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                return (model != null && model.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                    ? SW_ENABLE : SW_DISABLE;
            }
            catch { return SW_DISABLE; }
        }

        // === Update all Farsi notes across all sheets ===
        public void OnUpdateFarsiNotes()
        {
            if (!RequireLicense()) return;  // NEW

            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Open a drawing to use this command.", "Update Farsi Notes");
                    return;
                }

                var stats = UpdateAllFarsiNotes(model);

                MessageBox.Show(
                    $"Sheets scanned: {stats.Sheets}\r\n" +
                    $"Notes inspected: {stats.Inspected}\r\n" +
                    $"Farsi notes reshaped: {stats.Updated}\r\n" +
                    $"Skipped (non-Farsi/no change): {stats.Skipped}",
                    "Update Farsi Notes");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while updating notes:\r\n" + ex.Message, "Update Farsi Notes");
                Debug.WriteLine("OnUpdateFarsiNotes error: " + ex);
            }
        }

        public int OnUpdateFarsiNotesEnable()
        {
            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                return (model != null && model.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                    ? SW_ENABLE : SW_DISABLE;
            }
            catch { return SW_DISABLE; }
        }

        // NEW: Register command
        public void OnRegister()
        {
            try
            {
                LicensingUI.ShowRegistrationDialog(null);

                VerifiedLicense lic; string why;
                if (LicenseGate.IsActivated(out lic, out why))
                    _swApp?.SendMsgToUser2("SW2025RibbonAddin is ACTIVATED.",
                        (int)swMessageBoxIcon_e.swMbInformation,
                        (int)swMessageBoxBtn_e.swMbOk);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Registration UI error:\r\n" + ex.Message, AddinTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public int OnRegisterEnable() => SW_ENABLE;
        #endregion

        #region Intercept default Edit Text (only for our tagged notes)
        private int OnCommandOpenPreNotify(int command, int userCommand)
        {
            try
            {
                if (command != CMD_EDIT_TEXT) return 0;

                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING) return 0;

                var selMgr = (ISelectionMgr)model.SelectionManager;
                if (selMgr == null || selMgr.GetSelectedObjectCount2(-1) < 1) return 0;

                object selObj = selMgr.GetSelectedObject6(1, -1);
                INote note = null;

                if (selObj is INote n1) note = n1;
                else if (selObj is IAnnotation ann)
                {
                    try { note = (INote)ann.GetSpecificAnnotation(); } catch { }
                }

                if (note == null) return 0;

                // Only intercept notes created by this add-in (tag prefix)
                bool ours = false;
                try
                {
                    string tag = note.TagName;
                    ours = !string.IsNullOrEmpty(tag) && tag.StartsWith(FARSI_NOTE_TAG_PREFIX, StringComparison.Ordinal);
                }
                catch { }

                if (!ours) return 0;

                // Launch our editor instead of the default one
                if (!RequireLicense()) return 1;   // consume default, but do nothing if user cancels registration
                EditNoteWithFarsiEditor(note, model);
                return 1; // consume the default command
            }
            catch
            {
                return 0;
            }
        }
        #endregion

        #region Helpers  (unchanged from your file)
        private void StartFarsiNotePlacement(IModelDoc2 model, string preparedText, string fontName, double fontSizePts)
        {
            try
            {
                var view = (ModelView)model.GetFirstModelView();
                if (view == null)
                {
                    MessageBox.Show("Could not access the current view.", "Farsi Note");
                    return;
                }

                _activePlacement?.Dispose();
                _activePlacement = new FarsiNotePlacementSession(this, model, view, preparedText, fontName, fontSizePts);
                _activePlacement.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("StartFarsiNotePlacement: " + ex);
                MessageBox.Show("Could not start placement.\r\n" + ex.Message, "Farsi Note");
            }
        }

        private void EditNoteWithFarsiEditor(INote note, IModelDoc2 model)
        {
            string currentText = "";
            try { currentText = note.GetText(); }
            catch { currentText = ""; }

            string editable = ArabicNoteCodec.DecodeFromNote(currentText);

            IAnnotation ann = null;
            try { ann = (IAnnotation)note.GetAnnotation(); } catch { }

            ITextFormat tf = null;
            try { tf = ann != null ? (ITextFormat)ann.GetTextFormat(0) : null; } catch { }

            string fontName = tf?.TypeFaceName ?? "Tahoma";
            double sizePts = 12.0;
            try { sizePts = tf != null ? tf.CharHeightInPts : 12.0; } catch { }

            using (var dlg = new Forms.FarsiNoteForm())
            {
                dlg.NoteText = editable;
                dlg.SelectedFontName = fontName;
                dlg.FontSizePoints = sizePts;
                dlg.InsertJoiners = true;
                dlg.UseRtlMarkers = false;

                if (dlg.ShowDialog() != DialogResult.OK) return;

                string newRaw = dlg.NoteText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(newRaw)) return;

                string shaped = ArabicTextUtils.PrepareForSolidWorks(newRaw, dlg.UseRtlMarkers, dlg.InsertJoiners);

                try { note.SetText(shaped); } catch { return; }

                try
                {
                    ann = (IAnnotation)note.GetAnnotation();
                    tf = (ITextFormat)ann.GetTextFormat(0);
                    tf.TypeFaceName = dlg.SelectedFontName;
                    tf.CharHeightInPts = (int)Math.Round(dlg.FontSizePoints);
                    ann.SetTextFormat(0, false, tf);
                }
                catch { /* ignore */ }

                try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }
                try { model.GraphicsRedraw2(); } catch { }
            }
        }

        private struct UpdateStats { public int Sheets, Inspected, Updated, Skipped; }

        private UpdateStats UpdateAllFarsiNotes(IModelDoc2 model)
        {
            var stats = new UpdateStats();

            IDrawingDoc drw = model as IDrawingDoc;
            if (drw == null) return stats;

            // Remember current sheet
            string originalSheetName = null;
            try
            {
                var cur = (Sheet)drw.GetCurrentSheet();
                if (cur != null) originalSheetName = cur.GetName();
            }
            catch { }

            string[] sheetNames = GetSheetNamesSafe(drw);
            if (sheetNames == null || sheetNames.Length == 0) return stats;

            foreach (var sheetName in sheetNames)
            {
                try { drw.ActivateSheet(sheetName); } catch { }
                stats.Sheets++;

                IView v = drw.GetFirstView() as IView;
                while (v != null)
                {
                    try
                    {
                        INote note = v.GetFirstNote() as INote;
                        while (note != null)
                        {
                            stats.Inspected++;

                            string text = "";
                            try { text = note.GetText(); } catch { }

                            if (!string.IsNullOrEmpty(text) && LooksArabicOrPresentationForms(text))
                            {
                                string decoded = ArabicNoteCodec.DecodeFromNote(text);
                                string reshaped = ArabicTextUtils.PrepareForSolidWorks(decoded, false, true);

                                bool changed = !string.Equals(reshaped, text, StringComparison.Ordinal);
                                try { note.SetText(reshaped); } catch { }
                                try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }

                                if (changed) stats.Updated++; else stats.Skipped++;
                            }
                            else
                            {
                                stats.Skipped++;
                            }

                            object nextNoteObj = null;
                            try { nextNoteObj = note.GetNext(); } catch { }
                            note = nextNoteObj as INote;
                        }
                    }
                    catch { /* continue with next view */ }

                    object nextViewObj = null;
                    try { nextViewObj = v.GetNextView(); } catch { }
                    v = nextViewObj as IView;
                }

                try { model.GraphicsRedraw2(); } catch { }
            }

            // Restore original sheet
            if (!string.IsNullOrEmpty(originalSheetName))
            {
                try { drw.ActivateSheet(originalSheetName); } catch { }
            }

            return stats;
        }

        private static string[] GetSheetNamesSafe(IDrawingDoc drw)
        {
            try
            {
                var obj = drw.GetSheetNames();
                if (obj is string[] sarr) return sarr;
                if (obj is object[] oarr)
                {
                    var res = new string[oarr.Length];
                    for (int i = 0; i < oarr.Length; i++) res[i] = Convert.ToString(oarr[i]);
                    return res;
                }
            }
            catch { }
            return null;
        }

        private static bool LooksArabicOrPresentationForms(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;

            foreach (char ch in s)
            {
                int code = ch;
                // Arabic & Persian ranges + presentation forms
                if ((code >= 0x0600 && code <= 0x06FF) ||
                    (code >= 0x0750 && code <= 0x077F) ||
                    (code >= 0x08A0 && code <= 0x08FF) ||
                    (code >= 0xFB50 && code <= 0xFDFF) ||
                    (code >= 0xFE70 && code <= 0xFEFF))
                    return true;
            }
            return false;
        }
        #endregion

        #region COM Registration (unchanged)
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                var key = Registry.LocalMachine.CreateSubKey($@"Software\SolidWorks\Addins\{{{t.GUID}}}");
                key.SetValue(null, 1);
                key.SetValue("Title", AddinTitle);
                key.SetValue("Description", "Sample ribbon add-in with Farsi note support");
                key.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"COM registration failed: {ex.Message}\r\nTry running Visual Studio as administrator.", AddinTitle);
            }

            try
            {
                var keyCU = Registry.CurrentUser.CreateSubKey($@"Software\SolidWorks\AddinsStartup\{{{t.GUID}}}");
                keyCU.SetValue(null, 1);
                keyCU.Close();
            }
            catch { }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try { Registry.LocalMachine.DeleteSubKeyTree($@"Software\SolidWorks\Addins\{{{t.GUID}}}", false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree($@"Software\SolidWorks\AddinsStartup\{{{t.GUID}}}", false); } catch { }
        }
        #endregion

        #region Placement Session (unchanged)
        private sealed class FarsiNotePlacementSession : IDisposable
        {
            private readonly Addin _owner;
            private readonly IModelDoc2 _model;
            private readonly ModelView _view;
            private readonly string _text;
            private readonly string _fontName;
            private readonly double _fontSizePts;

            private SwMouse _mouse;
            private Forms.MouseGhostForm _ghost;
            private bool _active;

            public FarsiNotePlacementSession(Addin owner, IModelDoc2 model, ModelView view, string text, string fontName, double fontSizePts)
            {
                _owner = owner;
                _model = model;
                _view = view;
                _text = text;
                _fontName = string.IsNullOrWhiteSpace(fontName) ? "Tahoma" : fontName;
                _fontSizePts = fontSizePts <= 0 ? 12.0 : fontSizePts;
            }

            public void Start()
            {
                try
                {
                    _mouse = (SwMouse)_view.GetMouse();

                    // subscribe to events
                    _mouse.MouseMoveNotify += OnMouseMoveNotify;
                    _mouse.MouseLBtnDownNotify += OnMouseLBtnDownNotify;
                    _mouse.MouseRBtnDownNotify += OnMouseRBtnDownNotify;
                    _mouse.MouseSelectNotify += OnMouseSelectNotify;

                    _ghost = new Forms.MouseGhostForm();
                    _ghost.InitializeForText(_text, _fontName, (float)_fontSizePts);
                    _ghost.StartFollowing();
                    _active = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Placement Start failed: " + ex);
                    Cleanup();
                }
            }

            private int OnMouseMoveNotify(int x, int y, int WParam) => 0;
            private int OnMouseLBtnDownNotify(int x, int y, int WParam) => 0;

            private int OnMouseRBtnDownNotify(int x, int y, int WParam)
            {
                Cleanup();
                return 1; // consume
            }

            private int OnMouseSelectNotify(int Ix, int Iy, double x, double y, double z)
            {
                try
                {
                    var noteObj = _model.InsertNote(_text);
                    if (noteObj == null)
                    {
                        MessageBox.Show("Failed to insert note.", "Farsi Note");
                        Cleanup();
                        return 1;
                    }

                    var note = (INote)noteObj;
                    var ann = (IAnnotation)note.GetAnnotation();

                    try { note.TagName = $"{FARSI_NOTE_TAG_PREFIX}:{Guid.NewGuid():N}"; } catch { }

                    try
                    {
                        var tf = (ITextFormat)ann.GetTextFormat(0);
                        tf.TypeFaceName = _fontName;
                        tf.CharHeightInPts = (int)Math.Round(_fontSizePts);
                        ann.SetTextFormat(0, false, tf);
                    }
                    catch { /* ignore */ }

                    try { note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight); } catch { }

                    ann?.SetPosition2(x, y, 0.0);
                    _model.GraphicsRedraw2();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error placing note: " + ex.Message, "Farsi Note");
                    Debug.WriteLine("OnMouseSelectNotify error: " + ex);
                }
                finally
                {
                    Cleanup();
                }

                return 1;
            }

            public void Dispose() => Cleanup();

            private void Cleanup()
            {
                if (!_active) return;
                _active = false;

                try
                {
                    if (_mouse != null)
                    {
                        _mouse.MouseMoveNotify -= OnMouseMoveNotify;
                        _mouse.MouseLBtnDownNotify -= OnMouseLBtnDownNotify;
                        _mouse.MouseRBtnDownNotify -= OnMouseRBtnDownNotify;
                        _mouse.MouseSelectNotify -= OnMouseSelectNotify;
                    }
                }
                catch { /* ignore */ }

                try { _ghost?.Close(); _ghost?.Dispose(); } catch { }
                _ghost = null;
                _mouse = null;
            }
        }
        #endregion
    }
}
