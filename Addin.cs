using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SW2026RibbonAddin
{
    [ComVisible(true)]
    [Guid("B67E2D5A-8C73-4A3E-93B6-1761C1A8C0C5")]
    [ProgId("SW2026RibbonAddin.Addin")]
    public class Addin : ISwAddin
    {
        private SldWorks _swApp;
        private int _cookie;
        private ICommandManager _cmdMgr;
        private ICommandGroup _cmdGroup;

        // Command item indexes
        private int _helloCmdIndex = -1;
        private int _farsiCmdIndex = -1;
        private int _editSelNoteCmdIndex = -1;
        private int _updateAllNotesCmdIndex = -1;

        // Bump once if icons are cached and won’t refresh
        private const int MAIN_CMD_GROUP_ID = 20258;

        private const string TAB_NAME = "Mehdi";
        private const string MAIN_CMD_GROUP_TITLE = "Mehdi Tools";
        private const string MAIN_CMD_GROUP_TOOLTIP = "Custom tools";

        private const string HELLO_CMD_NAME = "Hello";
        private const string HELLO_CMD_TOOLTIP = "Show a hello message";
        private const string HELLO_CMD_HINT = "Hello";

        private const string FARSI_CMD_NAME = "Add Farsi Note";
        private const string FARSI_CMD_TOOLTIP = "Type Persian (Farsi) text and place it as a drawing note";
        private const string FARSI_CMD_HINT = "Farsi Note";

        private const string EDIT_SEL_NOTE_CMD_NAME = "Edit Selected Note (Farsi)";
        private const string EDIT_SEL_NOTE_CMD_TOOLTIP = "Open the Farsi editor for the selected note";
        private const string EDIT_SEL_NOTE_CMD_HINT = "Edit Note (Farsi)";

        private const string UPDATE_ALL_NOTES_CMD_NAME = "Update Farsi Notes";
        private const string UPDATE_ALL_NOTES_CMD_TOOLTIP = "Re‑shape and fix all Farsi notes across all sheets";
        private const string UPDATE_ALL_NOTES_CMD_HINT = "Update Farsi Notes";

        internal const string FARSI_NOTE_TAG_PREFIX = "MEHDI_FARSI_NOTE";

        private const int SW_ENABLE = 1;
        private const int SW_DISABLE = 0;
        private const int CMD_EDIT_TEXT = 1811;

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

                _swApp.CommandOpenPreNotify += OnCommandOpenPreNotify;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        public bool DisconnectFromSW()
        {
            try { if (_swApp != null) _swApp.CommandOpenPreNotify -= OnCommandOpenPreNotify; }
            catch { }
            try { if (_cmdMgr != null && _cmdGroup != null) _cmdMgr.RemoveCommandGroup(MAIN_CMD_GROUP_ID); }
            catch { }
            _cmdGroup = null;
            _cmdMgr = null;
            _swApp = null;
            return true;
        }
        #endregion

        #region UI (Command group + icons + tab without duplicates)
        private void CreateUI()
        {
            int errors = 0;

            // Remove any cached CommandGroup with same ID (avoids stale icons)
            try
            {
                var existing = _cmdMgr.GetCommandGroup(MAIN_CMD_GROUP_ID);
                if (existing != null) _cmdMgr.RemoveCommandGroup(MAIN_CMD_GROUP_ID);
            }
            catch { }

            const bool ignorePrevious = true;
            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                MAIN_CMD_GROUP_ID, MAIN_CMD_GROUP_TITLE,
                MAIN_CMD_GROUP_TOOLTIP, "", -1,
                ignorePrevious, ref errors);

            // ---- Assign transparent PNG strips (alpha) ----
            var (smallStrip, largeStrip) = EnsurePngStrips();
            if (File.Exists(smallStrip) && File.Exists(largeStrip))
            {
                try
                {
                    _cmdGroup.IconList = smallStrip;
                    _cmdGroup.MainIconList = largeStrip;
                    TrySetProperty(_cmdGroup, "SmallIconList", smallStrip);
                    TrySetProperty(_cmdGroup, "LargeIconList", largeStrip);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Assigning PNG strips failed: " + ex.Message);
                }
            }

            int itemOpts = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            _helloCmdIndex = _cmdGroup.AddCommandItem2(
                HELLO_CMD_NAME, -1, HELLO_CMD_TOOLTIP, HELLO_CMD_HINT,
                0, nameof(OnHello), nameof(OnHelloEnable), 1, itemOpts);

            _farsiCmdIndex = _cmdGroup.AddCommandItem2(
                FARSI_CMD_NAME, -1, FARSI_CMD_TOOLTIP, FARSI_CMD_HINT,
                1, nameof(OnAddFarsiNote), nameof(OnAddFarsiNoteEnable), 2, itemOpts);

            _editSelNoteCmdIndex = _cmdGroup.AddCommandItem2(
                EDIT_SEL_NOTE_CMD_NAME, -1, EDIT_SEL_NOTE_CMD_TOOLTIP, EDIT_SEL_NOTE_CMD_HINT,
                2, nameof(OnEditSelectedNoteFarsi), nameof(OnEditSelectedNoteFarsiEnable), 3, itemOpts);

            _updateAllNotesCmdIndex = _cmdGroup.AddCommandItem2(
                UPDATE_ALL_NOTES_CMD_NAME, -1, UPDATE_ALL_NOTES_CMD_TOOLTIP, UPDATE_ALL_NOTES_CMD_HINT,
                3, nameof(OnUpdateFarsiNotes), nameof(OnUpdateFarsiNotesEnable), 4, itemOpts);

            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();

            // ---- Command Tab (DRAWING) — remove/clear before creating to avoid duplicates ----
            try
            {
                int docType = (int)swDocumentTypes_e.swDocDRAWING;

                // Remove all existing tabs named "Mehdi" (prevents accumulation across reloads)
                RemoveAllTabsNamed(TAB_NAME);

                // Create a fresh tab
                var tab = _cmdMgr.AddCommandTab(docType, TAB_NAME);
                if (tab != null)
                {
                    var box = tab.AddCommandTabBox();

                    var ids = new int[]
                    {
                        _cmdGroup.get_CommandID(_helloCmdIndex),
                        _cmdGroup.get_CommandID(_farsiCmdIndex),
                        _cmdGroup.get_CommandID(_editSelNoteCmdIndex),
                        _cmdGroup.get_CommandID(_updateAllNotesCmdIndex)
                    };

                    var textTypes = new int[]
                    {
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
                Debug.WriteLine("Create tab failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Removes any existing command tabs that have the specified title.
        /// Works across PIA variants (ICommandTab vs CommandTab) using reflection.
        /// </summary>
        private void RemoveAllTabsNamed(string title)
        {
            try
            {
                int[] docTypes =
                {
                    (int)swDocumentTypes_e.swDocDRAWING,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swDocumentTypes_e.swDocASSEMBLY
                };

                foreach (int dt in docTypes)
                {
                    object tab = null;
                    try { tab = _cmdMgr.GetCommandTab(dt, title); }
                    catch { }

                    if (tab != null)
                    {
                        if (!TryRemoveCommandTab(tab))
                        {
                            // Fallback: clear pre‑existing boxes
                            TryClearTabBoxes(tab);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("RemoveAllTabsNamed: " + ex.Message);
            }
        }

        /// <summary>
        /// Attempts several RemoveCommandTab signatures by reflection.
        /// Returns true if a remove call succeeded.
        /// </summary>
        private bool TryRemoveCommandTab(object tab)
        {
            var cmType = _cmdMgr.GetType();

            // Common signature: RemoveCommandTab(CommandTab)
            foreach (var name in new[] { "RemoveCommandTab", "RemoveCommandTab2", "RemoveCommandTab3" })
            {
                try
                {
                    cmType.InvokeMember(name,
                        BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                        null, _cmdMgr, new object[] { tab });
                    return true;
                }
                catch (MissingMethodException) { }
                catch (TargetInvocationException) { }
                catch { }
            }

            // Some older PIAs have overloads RemoveCommandTab(int docType, string title)
            foreach (var name in new[] { "RemoveCommandTab", "RemoveCommandTab2", "RemoveCommandTab3" })
            {
                try
                {
                    int[] docTypes =
                    {
                        (int)swDocumentTypes_e.swDocDRAWING,
                        (int)swDocumentTypes_e.swDocPART,
                        (int)swDocumentTypes_e.swDocASSEMBLY
                    };

                    var titleProp = tab.GetType().GetProperty("Title", BindingFlags.Instance | BindingFlags.Public);
                    string title = titleProp != null ? Convert.ToString(titleProp.GetValue(tab)) : "Mehdi";

                    foreach (var dt in docTypes)
                    {
                        cmType.InvokeMember(name,
                            BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                            null, _cmdMgr, new object[] { dt, title });
                    }

                    return true;
                }
                catch (MissingMethodException) { }
                catch (TargetInvocationException) { }
                catch { }
            }

            return false;
        }

        /// <summary>
        /// If we cannot remove the tab, clear all existing boxes so we do not duplicate them.
        /// </summary>
        private void TryClearTabBoxes(object tab)
        {
            try
            {
                var getBox = tab.GetType().GetMethod("GetCommandTabBox",
                    BindingFlags.Instance | BindingFlags.Public);

                var getCount = tab.GetType().GetMethod("GetCommandTabBoxCount",
                    BindingFlags.Instance | BindingFlags.Public);

                var boxesProp = tab.GetType().GetProperty("CommandTabBoxes",
                    BindingFlags.Instance | BindingFlags.Public);

                if (getCount != null && getBox != null)
                {
                    int count = Convert.ToInt32(getCount.Invoke(tab, null));
                    for (int i = count - 1; i >= 0; i--)
                    {
                        var box = getBox.Invoke(tab, new object[] { i });
                        TryRemoveTabBox(tab, box);
                    }
                }
                else if (boxesProp != null)
                {
                    var boxes = boxesProp.GetValue(tab) as System.Collections.IEnumerable;
                    if (boxes != null)
                    {
                        foreach (var box in boxes)
                            TryRemoveTabBox(tab, box);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("TryClearTabBoxes: " + ex.Message);
            }
        }

        private void TryRemoveTabBox(object tab, object box)
        {
            if (box == null) return;

            try
            {
                var m = tab.GetType().GetMethod("RemoveCommandTabBox",
                    BindingFlags.Instance | BindingFlags.Public);
                if (m != null) m.Invoke(tab, new object[] { box });
            }
            catch { }
        }

        // Build PNG strips at %LOCALAPPDATA%\SW2026RibbonAddin\icons\ from embedded or output files
        private (string small, string large) EnsurePngStrips()
        {
            string outDir = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "SW2026RibbonAddin", "icons");
            Directory.CreateDirectory(outDir);

            string small = Path.Combine(outDir, "small_strip.png");
            string large = Path.Combine(outDir, "large_strip.png");

            try
            {
                string[] smallFiles = { "hello_20.png", "farsi_20.png", "edit_20.png", "update_20.png" };
                string[] largeFiles = { "hello_32.png", "farsi_32.png", "edit_32.png", "update_32.png" };

                BuildStrip(ResolveImages(smallFiles), 20, small);
                BuildStrip(ResolveImages(largeFiles), 32, large);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EnsurePngStrips failed: " + ex.Message);
            }

            return (small, large);
        }

        // Tries: (1) embedded resource, (2) Resources folder next to the DLL
        private Stream[] ResolveImages(string[] fileNames)
        {
            var asm = Assembly.GetExecutingAssembly();
            var allNames = asm.GetManifestResourceNames();

            Stream ResolveOne(string fname)
            {
                string hit = allNames.FirstOrDefault(n =>
                    n.EndsWith(".Resources." + fname, StringComparison.OrdinalIgnoreCase));

                if (hit != null) return asm.GetManifestResourceStream(hit);

                string asmDir = Path.GetDirectoryName(asm.Location) ?? ".";
                string candidate = Path.Combine(asmDir, "Resources", fname);
                if (File.Exists(candidate)) return File.OpenRead(candidate);

                throw new FileNotFoundException("Icon not found: " + fname);
            }

            return fileNames.Select(ResolveOne).ToArray();
        }

        private void BuildStrip(Stream[] images, int size, string outPng)
        {
            int W = size * images.Length;
            using (var strip = new Bitmap(W, size, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(strip))
            {
                g.Clear(Color.Transparent);
                int x = 0;

                foreach (var s in images)
                {
                    using (s)
                    using (var img = Image.FromStream(s))
                    {
                        g.DrawImage(img, new Rectangle(x, 0, size, size));
                        x += size;
                    }
                }

                strip.Save(outPng, ImageFormat.Png);
            }
        }

        private static void TrySetProperty(object obj, string prop, object value)
        {
            try
            {
                var pi = obj.GetType().GetProperty(prop, BindingFlags.Instance | BindingFlags.Public);
                if (pi != null && pi.CanWrite) pi.SetValue(obj, value);
            }
            catch { }
        }
        #endregion

        #region Commands
        public void OnHello()
        {
            try
            {
                MessageBox.Show("Hello from Mehdi Tools ✨", "SW2026RibbonAddin");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        public int OnHelloEnable() => SW_ENABLE;

        public void OnAddFarsiNote()
        {
            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
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

                    StartFarsiNotePlacement(
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

        public int OnAddFarsiNoteEnable()
        {
            try
            {
                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                return (model != null && model.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                    ? SW_ENABLE
                    : SW_DISABLE;
            }
            catch
            {
                return SW_DISABLE;
            }
        }

        public void OnEditSelectedNoteFarsi()
        {
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

                if (selObj is INote n1)
                    note = n1;
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
                    ? SW_ENABLE
                    : SW_DISABLE;
            }
            catch
            {
                return SW_DISABLE;
            }
        }

        public void OnUpdateFarsiNotes()
        {
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
                    $"Skipped (non‑Farsi/no change): {stats.Skipped}",
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
                    ? SW_ENABLE
                    : SW_DISABLE;
            }
            catch
            {
                return SW_DISABLE;
            }
        }
        #endregion

        #region Intercept default Edit Text for our tagged notes
        private int OnCommandOpenPreNotify(int command, int userCommand)
        {
            try
            {
                if (command != CMD_EDIT_TEXT) return 0;

                var model = _swApp?.IActiveDoc2 as IModelDoc2;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                    return 0;

                var selMgr = (ISelectionMgr)model.SelectionManager;
                if (selMgr == null || selMgr.GetSelectedObjectCount2(-1) < 1)
                    return 0;

                object selObj = selMgr.GetSelectedObject6(1, -1);
                INote note = null;

                if (selObj is INote n1)
                    note = n1;
                else if (selObj is IAnnotation ann)
                {
                    try { note = (INote)ann.GetSpecificAnnotation(); } catch { }
                }

                if (note == null)
                    return 0;

                bool ours = false;
                try
                {
                    string tag = note.TagName;
                    ours = !string.IsNullOrEmpty(tag) &&
                        tag.StartsWith(FARSI_NOTE_TAG_PREFIX, StringComparison.Ordinal);
                }
                catch { }

                if (!ours) return 0;

                EditNoteWithFarsiEditor(note, model);
                return 1; // consume default command
            }
            catch
            {
                return 0;
            }
        }
        #endregion

        #region Helpers (placement + editor + batch update)
        private void StartFarsiNotePlacement(
            IModelDoc2 model,
            string preparedText,
            string fontName,
            double fontSizePts,
            HorizontalAlignment alignment)
        {
            try
            {
                var view = (ModelView)model.GetFirstModelView();
                if (view == null)
                {
                    MessageBox.Show("Could not access the current view.", "Farsi Note");
                    return;
                }

                int justification = MapAlignmentToSwJustification(alignment);

                _activePlacement?.Dispose();
                _activePlacement = new FarsiNotePlacementSession(
                    this, model, view,
                    preparedText, fontName, fontSizePts,
                    justification);
                _activePlacement.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("StartFarsiNotePlacement: " + ex);
                MessageBox.Show("Could not start placement.\r\n" + ex.Message, "Farsi Note");
            }
        }

        private static int MapAlignmentToSwJustification(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Left:
                    return (int)swTextJustification_e.swTextJustificationLeft;
                case HorizontalAlignment.Center:
                    return (int)swTextJustification_e.swTextJustificationCenter;
                case HorizontalAlignment.Right:
                default:
                    return (int)swTextJustification_e.swTextJustificationRight;
            }
        }

        private void EditNoteWithFarsiEditor(INote note, IModelDoc2 model)
        {
            string currentText = "";
            try { currentText = note.GetText(); }
            catch { currentText = ""; }

            string editable = ArabicNoteCodec.DecodeFromNote(currentText);

            // Edit-mode: text only (no formatting changes)
            using (var dlg = new Forms.FarsiNoteForm(false))
            {
                dlg.NoteText = editable;
                dlg.InsertJoiners = true;
                dlg.UseRtlMarkers = false;

                if (dlg.ShowDialog() != DialogResult.OK) return;

                string newRaw = dlg.NoteText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(newRaw)) return;

                string shaped = ArabicTextUtils.PrepareForSolidWorks(
                    newRaw, dlg.UseRtlMarkers, dlg.InsertJoiners);

                try { note.SetText(shaped); }
                catch { return; }

                try { model.GraphicsRedraw2(); }
                catch { }
            }
        }

        private struct UpdateStats
        {
            public int Sheets;
            public int Inspected;
            public int Updated;
            public int Skipped;
        }

        private UpdateStats UpdateAllFarsiNotes(IModelDoc2 model)
        {
            var stats = new UpdateStats();

            IDrawingDoc drw = model as IDrawingDoc;
            if (drw == null) return stats;

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
                try { drw.ActivateSheet(sheetName); }
                catch { }

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
                            try { text = note.GetText(); }
                            catch { }

                            if (!string.IsNullOrEmpty(text) && LooksArabicOrPresentationForms(text))
                            {
                                string decoded = ArabicNoteCodec.DecodeFromNote(text);
                                string reshaped = ArabicTextUtils.PrepareForSolidWorks(decoded, false, true);

                                bool changed = !string.Equals(reshaped, text, StringComparison.Ordinal);
                                try { note.SetText(reshaped); }
                                catch { }

                                if (changed) stats.Updated++;
                                else stats.Skipped++;
                            }
                            else
                            {
                                stats.Skipped++;
                            }

                            object nextNoteObj = null;
                            try { nextNoteObj = note.GetNext(); }
                            catch { }

                            note = nextNoteObj as INote;
                        }
                    }
                    catch
                    {
                        // continue
                    }

                    object nextViewObj = null;
                    try { nextViewObj = v.GetNextView(); }
                    catch { }

                    v = nextViewObj as IView;
                }

                try { model.GraphicsRedraw2(); }
                catch { }
            }

            if (!string.IsNullOrEmpty(originalSheetName))
            {
                try { drw.ActivateSheet(originalSheetName); }
                catch { }
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
                    for (int i = 0; i < oarr.Length; i++)
                        res[i] = Convert.ToString(oarr[i]);
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

        #region COM Registration
        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                var key = Registry.LocalMachine.CreateSubKey($@"Software\SolidWorks\Addins\{{{t.GUID}}}");
                key.SetValue(null, 1);
                key.SetValue("Title", "SW2026RibbonAddin");
                key.SetValue("Description", "Sample ribbon add-in with Farsi note support");
                key.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"COM registration failed: {ex.Message}\r\nTry running Visual Studio as administrator.",
                    "SW2026RibbonAddin");
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
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(
                    $@"Software\SolidWorks\Addins\{{{t.GUID}}}", false);
            }
            catch { }

            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(
                    $@"Software\SolidWorks\AddinsStartup\{{{t.GUID}}}", false);
            }
            catch { }
        }
        #endregion

        #region Placement Session (mouse + overlay)
        private sealed class FarsiNotePlacementSession : IDisposable
        {
            private readonly Addin _owner;
            private readonly IModelDoc2 _model;
            private readonly ModelView _view;
            private readonly string _text;
            private readonly string _fontName;
            private readonly double _fontSizePts;
            private readonly int _justification;

            private SolidWorks.Interop.sldworks.Mouse _mouse;
            private Forms.MouseGhostForm _ghost;
            private bool _active;

            public FarsiNotePlacementSession(
                Addin owner,
                IModelDoc2 model,
                ModelView view,
                string text,
                string fontName,
                double fontSizePts,
                int justification)
            {
                _owner = owner;
                _model = model;
                _view = view;
                _text = text;
                _fontName = string.IsNullOrWhiteSpace(fontName) ? "Tahoma" : fontName;
                _fontSizePts = fontSizePts <= 0 ? 12.0 : fontSizePts;
                _justification = justification;
            }

            public void Start()
            {
                try
                {
                    _mouse = (SolidWorks.Interop.sldworks.Mouse)_view.GetMouse();

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
                return 1;
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

                    try
                    {
                        note.TagName = $"{FARSI_NOTE_TAG_PREFIX}:{Guid.NewGuid():N}";
                    }
                    catch { }

                    try
                    {
                        var tf = (ITextFormat)ann.GetTextFormat(0);
                        tf.TypeFaceName = _fontName;
                        tf.CharHeightInPts = (int)Math.Round(_fontSizePts);
                        ann.SetTextFormat(0, false, tf);
                    }
                    catch { }

                    try
                    {
                        note.SetTextJustification(_justification);
                    }
                    catch { }

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
                catch { }

                try
                {
                    _ghost?.Close();
                    _ghost?.Dispose();
                }
                catch { }

                _ghost = null;
                _mouse = null;
            }
        }
        #endregion
    }
}
