using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SW2026RibbonAddin.Commands;
using SW2026RibbonAddin.Licensing;
using Forms = SW2026RibbonAddin.Forms;

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

        // Button infrastructure
        private const int MAX_CMD_SLOTS = 16;
        private readonly List<IMehdiRibbonButton> _buttons = new List<IMehdiRibbonButton>();
        private readonly IMehdiRibbonButton[] _slotButtons = new IMehdiRibbonButton[MAX_CMD_SLOTS];
        private readonly int[] _slotCommandIndices = new int[MAX_CMD_SLOTS];

        private LicenseState _licenseState = LicenseState.TrialActive;

        // Legacy indices (no longer used, but kept for compatibility)
        private int _helloCmdIndex = -1;
        private int _farsiCmdIndex = -1;
        private int _editSelNoteCmdIndex = -1;
        private int _updateAllNotesCmdIndex = -1;

        // Command group + tab constants
        private const int MAIN_CMD_GROUP_ID = 20258;
        private const string TAB_NAME = "Mehdi";
        private const string MAIN_CMD_GROUP_TITLE = "Mehdi Tools";
        private const string MAIN_CMD_GROUP_TOOLTIP = "Custom tools";

        // Top‑level classic menu (root “Mehdi” menu)
        private const string MEHDI_MENU_CAPTION = "Mehdi";
        // 5 = between Tools and Window, per usual SolidWorks examples :contentReference[oaicite:0]{index=0}
        private const int MEHDI_MENU_POSITION = 5;

        // Farsi note command strings (legacy text)
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
        private const string UPDATE_ALL_NOTES_CMD_TOOLTIP = "Re-shape and fix all Farsi notes across all sheets";
        private const string UPDATE_ALL_NOTES_CMD_HINT = "Update Farsi Notes";

        internal const string FARSI_NOTE_TAG_PREFIX = "MEHDI_FARSI_NOTE";

        private const int SW_ENABLE = 1;
        private const int SW_DISABLE = 0;

        // SolidWorks command ID for "Edit Text" on a note
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

                InitializeLicenseState();
                InitializeButtons();
                CreateUI();

                _swApp.CommandOpenPreNotify += OnCommandOpenPreNotify;

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ConnectToSW failed: " + ex);
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
                if (_cmdMgr != null && _cmdGroup != null)
                    _cmdMgr.RemoveCommandGroup(MAIN_CMD_GROUP_ID);
            }
            catch { }

            _cmdGroup = null;
            _cmdMgr = null;
            _swApp = null;

            return true;
        }

        #endregion

        #region Licensing + button discovery

        private void InitializeLicenseState()
        {
            try
            {
                if (LicenseGate.IsLicensed)
                    _licenseState = LicenseState.Licensed;
                else
                    _licenseState = LicenseState.TrialActive;
            }
            catch
            {
                _licenseState = LicenseState.TrialActive;
            }
        }

        private void InitializeButtons()
        {
            _buttons.Clear();

            try
            {
                var buttonType = typeof(IMehdiRibbonButton);
                var asm = Assembly.GetExecutingAssembly();

                var types = asm.GetTypes()
                    .Where(t =>
                        !t.IsAbstract &&
                        buttonType.IsAssignableFrom(t) &&
                        t.GetConstructor(Type.EmptyTypes) != null);

                foreach (var t in types)
                {
                    if (Activator.CreateInstance(t) is IMehdiRibbonButton button)
                        _buttons.Add(button);
                }

                // Fallback if reflection fails
                if (_buttons.Count == 0)
                {
                    _buttons.Add(new HelloButton());
                    _buttons.Add(new AddFarsiNoteButton());
                    _buttons.Add(new EditSelectedFarsiNoteButton());
                    _buttons.Add(new UpdateFarsiNotesButton());
                    _buttons.Add(new DwgButton());
                }

                // Sort by section, then section order, then display name
                _buttons.Sort((a, b) =>
                {
                    int sectionCompare = a.Section.CompareTo(b.Section);
                    if (sectionCompare != 0) return sectionCompare;

                    int orderCompare = a.SectionOrder.CompareTo(b.SectionOrder);
                    if (orderCompare != 0) return orderCompare;

                    return string.Compare(a.DisplayName, b.DisplayName,
                        StringComparison.CurrentCultureIgnoreCase);
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine("InitializeButtons failed: " + ex);
            }
        }

        #endregion

        #region UI creation (command group + tabs)

        private void CreateUI()
        {
            int errors = 0;

            // Remove previous command group with same ID (avoid stale icons)
            try
            {
                var existing = _cmdMgr.GetCommandGroup(MAIN_CMD_GROUP_ID);
                if (existing != null)
                    _cmdMgr.RemoveCommandGroup(MAIN_CMD_GROUP_ID);
            }
            catch { }

            const bool ignorePrevious = true;

            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                MAIN_CMD_GROUP_ID,
                MAIN_CMD_GROUP_TITLE,
                MAIN_CMD_GROUP_TOOLTIP,
                "",
                -1,
                ignorePrevious,
                ref errors);

            // Icon strips
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

            Array.Clear(_slotButtons, 0, _slotButtons.Length);
            Array.Clear(_slotCommandIndices, 0, _slotCommandIndices.Length);

            int slot = 0;

            foreach (var button in _buttons)
            {
                if (slot >= MAX_CMD_SLOTS)
                {
                    Debug.WriteLine("Max command slots exceeded; ignoring " + button.Id);
                    break;
                }

                string callback = $"OnCmd{slot}";
                string enable = $"OnCmd{slot}Enable";

                int cmdIndex = _cmdGroup.AddCommandItem2(
                    button.DisplayName,
                    -1,
                    button.Tooltip,
                    button.Hint,
                    slot,
                    callback,
                    enable,
                    slot + 1,
                    itemOpts);

                _slotButtons[slot] = button;
                _slotCommandIndices[slot] = cmdIndex;

                slot++;
            }

            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();

            // ----- Command tabs for PART / ASSEMBLY / DRAWING -----
            try
            {
                // Remove all existing "Mehdi" tabs first
                RemoveAllTabsNamed(TAB_NAME);

                int[] docTypes =
                {
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swDocumentTypes_e.swDocDRAWING
                };

                foreach (int docType in docTypes)
                {
                    var tab = _cmdMgr.AddCommandTab(docType, TAB_NAME);
                    if (tab == null) continue;

                    var groupsBySection = _buttons
                        .GroupBy(b => b.Section)
                        .OrderBy(g => g.Key);

                    foreach (var group in groupsBySection)
                    {
                        var box = tab.AddCommandTabBox();

                        var ids = group
                            .Select(b =>
                            {
                                int s = Array.IndexOf(_slotButtons, b);
                                if (s < 0) return -1;
                                int cidx = _slotCommandIndices[s];
                                return _cmdGroup.get_CommandID(cidx);
                            })
                            .Where(id => id != -1)
                            .ToArray();

                        if (ids.Length == 0) continue;

                        var textTypes = new int[ids.Length];
                        for (int i = 0; i < textTypes.Length; i++)
                            textTypes[i] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                        box.AddCommands(ids, textTypes);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CreateUI - tab creation failed: " + ex.Message);
            }
        }

        private void TrySetProperty(object target, string propertyName, object value)
        {
            try
            {
                var type = target.GetType();
                var prop = type.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(target, value, null);
                    return;
                }
            }
            catch { }

            try
            {
                var mi = target.GetType().GetMethod(propertyName,
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (mi != null)
                    mi.Invoke(target, new object[] { value });
            }
            catch { }
        }

        private (string small, string large) EnsurePngStrips()
        {
            string outDir = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "SW2026RibbonAddin",
                "icons");

            Directory.CreateDirectory(outDir);

            string small = Path.Combine(outDir, "small_strip.png");
            string large = Path.Combine(outDir, "large_strip.png");

            try
            {
                if (_buttons != null && _buttons.Count > 0)
                {
                    string[] smallFiles = _buttons.Select(b => b.SmallIconFile).ToArray();
                    string[] largeFiles = _buttons.Select(b => b.LargeIconFile).ToArray();

                    BuildStrip(ResolveImages(smallFiles), 20, small);
                    BuildStrip(ResolveImages(largeFiles), 32, large);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("EnsurePngStrips failed: " + ex.Message);
            }

            return (small, large);
        }

        private Stream[] ResolveImages(string[] fileNames)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resources = asm.GetManifestResourceNames();

            Stream ResolveOne(string fname)
            {
                string hit = resources.FirstOrDefault(n =>
                    n.EndsWith(".Resources." + fname, StringComparison.OrdinalIgnoreCase));

                if (hit != null)
                    return asm.GetManifestResourceStream(hit);

                string dir = Path.Combine(
                    Path.GetDirectoryName(asm.Location) ?? "",
                    "Resources");

                string path = Path.Combine(dir, fname);
                if (File.Exists(path))
                    return File.OpenRead(path);

                return null;
            }

            return fileNames.Select(ResolveOne).ToArray();
        }

        private void BuildStrip(Stream[] images, int iconSize, string outFile)
        {
            if (images == null || images.Length == 0)
                return;

            try
            {
                using (var first = images[0])
                {
                    if (first == null) return;

                    using (var bmpFirst = new Bitmap(first))
                    {
                        int width = images.Length * iconSize;
                        int height = iconSize;

                        using (var strip = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                        using (var g = Graphics.FromImage(strip))
                        {
                            g.Clear(Color.Transparent);

                            for (int i = 0; i < images.Length; i++)
                            {
                                var s = images[i];
                                if (s == null) continue;

                                using (var bmp = new Bitmap(s))
                                {
                                    g.DrawImage(
                                        bmp,
                                        new Rectangle(i * iconSize, 0, iconSize, iconSize),
                                        new Rectangle(0, 0, bmp.Width, bmp.Height),
                                        GraphicsUnit.Pixel);
                                }
                            }

                            strip.Save(outFile, ImageFormat.Png);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("BuildStrip failed: " + ex.Message);
            }
            finally
            {
                if (images != null)
                {
                    foreach (var s in images)
                    {
                        try { s?.Dispose(); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Remove all existing tabs named <paramref name="title"/> for Part, Assembly, and Drawing.
        /// Prevents multiple "Mehdi" tabs from accumulating.
        /// </summary>
        private void RemoveAllTabsNamed(string title)
        {
            if (_cmdMgr == null)
                return;

            try
            {
                int[] docTypes =
                {
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swDocumentTypes_e.swDocDRAWING
                };

                foreach (int dt in docTypes)
                {
                    while (true)
                    {
                        CommandTab tab = null;

                        try
                        {
                            tab = _cmdMgr.GetCommandTab(dt, title);
                        }
                        catch
                        {
                            tab = null;
                        }

                        if (tab == null)
                            break;

                        try
                        {
                            _cmdMgr.RemoveCommandTab(tab);
                        }
                        catch
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("RemoveAllTabsNamed failed: " + ex.Message);
            }
        }

        #endregion

        #region Command slot dispatch

        private void RunButtonSlot(int slotIndex)
        {
            try
            {
                if (slotIndex < 0 || slotIndex >= _slotButtons.Length)
                    return;

                var button = _slotButtons[slotIndex];
                if (button == null)
                    return;

                var context = new AddinContext(this, _swApp);

                if (button.IsFreeFeature)
                {
                    button.Execute(context);
                    return;
                }

                switch (_licenseState)
                {
                    case LicenseState.Licensed:
                    case LicenseState.TrialActive:
                        button.Execute(context);
                        break;

                    case LicenseState.TrialExpired:
                    case LicenseState.Unlicensed:
                        // For now we still run; hook trial-expired UI here later if needed
                        button.Execute(context);
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("RunButtonSlot error: " + ex);
            }
        }

        private int GetButtonEnableSlot(int slotIndex)
        {
            try
            {
                if (slotIndex < 0 || slotIndex >= _slotButtons.Length)
                    return SW_DISABLE;

                var button = _slotButtons[slotIndex];
                if (button == null)
                    return SW_DISABLE;

                var context = new AddinContext(this, _swApp);
                return button.GetEnableState(context);
            }
            catch
            {
                return SW_DISABLE;
            }
        }

        public void OnCmd0() => RunButtonSlot(0);
        public int OnCmd0Enable() => GetButtonEnableSlot(0);

        public void OnCmd1() => RunButtonSlot(1);
        public int OnCmd1Enable() => GetButtonEnableSlot(1);

        public void OnCmd2() => RunButtonSlot(2);
        public int OnCmd2Enable() => GetButtonEnableSlot(2);

        public void OnCmd3() => RunButtonSlot(3);
        public int OnCmd3Enable() => GetButtonEnableSlot(3);

        public void OnCmd4() => RunButtonSlot(4);
        public int OnCmd4Enable() => GetButtonEnableSlot(4);

        public void OnCmd5() => RunButtonSlot(5);
        public int OnCmd5Enable() => GetButtonEnableSlot(5);

        public void OnCmd6() => RunButtonSlot(6);
        public int OnCmd6Enable() => GetButtonEnableSlot(6);

        public void OnCmd7() => RunButtonSlot(7);
        public int OnCmd7Enable() => GetButtonEnableSlot(7);

        public void OnCmd8() => RunButtonSlot(8);
        public int OnCmd8Enable() => GetButtonEnableSlot(8);

        public void OnCmd9() => RunButtonSlot(9);
        public int OnCmd9Enable() => GetButtonEnableSlot(9);

        public void OnCmd10() => RunButtonSlot(10);
        public int OnCmd10Enable() => GetButtonEnableSlot(10);

        public void OnCmd11() => RunButtonSlot(11);
        public int OnCmd11Enable() => GetButtonEnableSlot(11);

        public void OnCmd12() => RunButtonSlot(12);
        public int OnCmd12Enable() => GetButtonEnableSlot(12);

        public void OnCmd13() => RunButtonSlot(13);
        public int OnCmd13Enable() => GetButtonEnableSlot(13);

        public void OnCmd14() => RunButtonSlot(14);
        public int OnCmd14Enable() => GetButtonEnableSlot(14);

        public void OnCmd15() => RunButtonSlot(15);
        public int OnCmd15Enable() => GetButtonEnableSlot(15);

        #endregion

        #region Legacy handlers (Farsi tools / Hello)

        public void OnHello()
        {
            try
            {
                MessageBox.Show("Hello from Mehdi Tools", "SW2026RibbonAddin");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("OnHello error: " + ex);
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

                using (var dlg = new Forms.FarsiNoteForm(true))
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return;

                    string text = dlg.NoteText ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(text))
                        return;

                    text = ArabicTextUtils.PrepareForSolidWorks(
                        text,
                        dlg.UseRtlMarkers,
                        dlg.InsertJoiners);

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
                Debug.WriteLine("OnAddFarsiNote error: " + ex);
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
                    ? SW_ENABLE
                    : SW_DISABLE;
            }
            catch
            {
                return SW_DISABLE;
            }
        }

        #endregion

        #region Intercept default Edit Text for tagged notes

        private int OnCommandOpenPreNotify(int command, int userCommand)
        {
            try
            {
                if (command != CMD_EDIT_TEXT)
                    return 0;

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

                if (!ours)
                    return 0;

                EditNoteWithFarsiEditor(note, model);
                return 1;   // consume default command
            }
            catch
            {
                return 0;
            }
        }

        #endregion

        #region Farsi note helpers

        internal void StartFarsiNotePlacement(
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

                int just = MapAlignmentToSwJustification(alignment);

                _activePlacement?.Dispose();
                _activePlacement = new FarsiNotePlacementSession(
                    this,
                    model,
                    view,
                    preparedText,
                    fontName,
                    fontSizePts,
                    just);

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

        internal void EditNoteWithFarsiEditor(INote note, IModelDoc2 model)
        {
            string currentText = "";
            try { currentText = note.GetText(); }
            catch { }

            string editable = ArabicNoteCodec.DecodeFromNote(currentText);

            using (var dlg = new Forms.FarsiNoteForm(false))
            {
                dlg.NoteText = editable;
                dlg.InsertJoiners = true;
                dlg.UseRtlMarkers = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                string newRaw = dlg.NoteText ?? string.Empty;
                if (string.IsNullOrWhiteSpace(newRaw))
                    return;

                string shaped = ArabicTextUtils.PrepareForSolidWorks(
                    newRaw,
                    dlg.UseRtlMarkers,
                    dlg.InsertJoiners);

                try { note.SetText(shaped); }
                catch { return; }

                try { model.GraphicsRedraw2(); }
                catch { }
            }
        }

        internal struct UpdateStats
        {
            public int Sheets;
            public int Inspected;
            public int Updated;
            public int Skipped;
        }

        internal UpdateStats UpdateAllFarsiNotes(IModelDoc2 model)
        {
            var stats = new UpdateStats();

            IDrawingDoc drw = model as IDrawingDoc;
            if (drw == null)
                return stats;

            string originalSheetName = null;
            try
            {
                var cur = (Sheet)drw.GetCurrentSheet();
                if (cur != null) originalSheetName = cur.GetName();
            }
            catch { }

            string[] sheetNames = GetSheetNamesSafe(drw);
            if (sheetNames == null || sheetNames.Length == 0)
                return stats;

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
                                string reshaped = ArabicTextUtils.PrepareForSolidWorks(
                                    decoded,
                                    false,
                                    true);

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
                            try { nextNoteObj = note.GetNext(); } catch { }
                            note = nextNoteObj as INote;
                        }
                    }
                    catch { }

                    object nextViewObj = null;
                    try { nextViewObj = v.GetNextView(); } catch { }
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
                if ((code >= 0x0600 && code <= 0x06FF) ||  // Arabic
                    (code >= 0x0750 && code <= 0x077F) ||  // Arabic Supplement
                    (code >= 0x08A0 && code <= 0x08FF) ||  // Arabic Extended-A
                    (code >= 0xFB50 && code <= 0xFDFF) ||  // Presentation Forms-A
                    (code >= 0xFE70 && code <= 0xFEFF))    // Presentation Forms-B
                    return true;
            }

            return false;
        }

        #endregion

        #region COM registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                var key = Registry.LocalMachine.CreateSubKey(
                    $@"Software\SolidWorks\Addins\{{{t.GUID}}}");

                key.SetValue(null, 1);
                key.SetValue("Title", "SW2026RibbonAddin");
                key.SetValue("Description", "Sample ribbon add-in with Farsi note and DWG tools");
                key.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"COM registration failed: {ex.Message}\r\n" +
                    "Try running Visual Studio as administrator.",
                    "SW2026RibbonAddin");
            }

            try
            {
                var keyCU = Registry.CurrentUser.CreateSubKey(
                    $@"Software\SolidWorks\AddinsStartup\{{{t.GUID}}}");

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

        #region Placement session (mouse + overlay)

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

                    try
                    {
                        ann?.SetPosition2(x, y, 0.0);
                    }
                    catch { }

                    try
                    {
                        _model.GraphicsRedraw2();
                    }
                    catch { }
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
