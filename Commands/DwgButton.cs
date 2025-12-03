using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SW2026RibbonAddin;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// DWG export button:
    /// - For a sheet-metal part: exports one DWG (geometry-only) into a new folder.
    /// - For an assembly: exports one DWG per unique sheet-metal part (recursively).
    /// - Uses a Windows "Open" style dialog to pick a MAIN folder, then creates a
    ///   subfolder named after the part/assembly (must NOT already exist).
    /// - Writes a CSV inside that subfolder with: FileName,PlateThickness_mm,Quantity.
    /// </summary>
    internal sealed class DwgButton : IMehdiRibbonButton
    {
        public string Id => "DWG";

        public string DisplayName => "DWG";
        public string Tooltip => "Export DWG files for sheet-metal parts";
        public string Hint => "DWG export for sheet metal";

        public string SmallIconFile => "dwg_20.png";
        public string LargeIconFile => "dwg_32.png";

        public RibbonSection Section => RibbonSection.General;
        public int SectionOrder => 10;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            var swApp = context.SwApp;
            var model = context.ActiveModel;

            if (model == null)
            {
                MessageBox.Show("No active document. Open a part or assembly first.", "DWG Export");
                return;
            }

            int docType = model.GetType();
            bool isPart = docType == (int)swDocumentTypes_e.swDocPART;
            bool isAsm = docType == (int)swDocumentTypes_e.swDocASSEMBLY;

            if (!isPart && !isAsm)
            {
                MessageBox.Show("DWG export is only available for parts and assemblies.", "DWG Export");
                return;
            }

            // 2) Let user pick the MAIN folder (standard file-style dialog)
            string mainFolder = SelectMainFolder(model);
            if (string.IsNullOrEmpty(mainFolder))
                return; // user cancelled

            // 3) Create subfolder named after the active document
            string baseName = GetDocumentBaseName(model);
            string jobFolder = Path.Combine(mainFolder, baseName);

            if (Directory.Exists(jobFolder))
            {
                MessageBox.Show(
                    "Output folder already exists:\r\n" + jobFolder +
                    "\r\n\r\nPlease delete it or choose another main folder.",
                    "DWG Export");
                return;
            }

            try
            {
                Directory.CreateDirectory(jobFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not create output folder:\r\n" + jobFolder +
                    "\r\n\r\n" + ex.Message,
                    "DWG Export");
                return;
            }

            // Remember original active doc so we can restore it
            string originalKey = GetActiveDocKey(swApp);

            try
            {
                if (isPart)
                {
                    RunForSinglePart(swApp, model, jobFolder);
                }
                else
                {
                    RunForAssembly(swApp, (IAssemblyDoc)model, jobFolder);
                }
            }
            finally
            {
                RestoreActiveDoc(swApp, originalKey);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // New strategy:
            // - As long as there is some active document, keep the DWG button enabled.
            // - We no longer try to inspect the part/assembly here.
            // - When the user clicks, Execute() will check document type and
            //   presence of sheet-metal and show a message if nothing can be done.

            try
            {
                var model = context.ActiveModel;
                return model != null ? AddinContext.Enable : AddinContext.Disable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }

        // ------------------------------------------------------------------
        //  Single part
        // ------------------------------------------------------------------

        private void RunForSinglePart(ISldWorks swApp, IModelDoc2 model, string jobFolder)
        {
            if (!TryGetSheetMetalThickness(model, out double thicknessMeters))
            {
                MessageBox.Show("The active part is not a sheet-metal part.", "DWG Export");
                return;
            }

            string modelPath = model.GetPathName();
            if (string.IsNullOrEmpty(modelPath))
            {
                MessageBox.Show(
                    "Save the part before exporting to DWG.",
                    "DWG Export");
                return;
            }

            var csvLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity"
            };

            int exported = 0;
            int failed = 0;

            // Part is already active, no need to re-activate
            if (ExportSinglePartToDwg(model, modelPath, jobFolder, out string dwgFileName))
            {
                exported++;
                double thicknessMm = thicknessMeters * 1000.0;
                csvLines.Add(
                    $"{dwgFileName},{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)},1");
            }
            else
            {
                failed++;
            }

            string csvPath = Path.Combine(jobFolder, "parts.csv");
            TryWriteCsv(csvPath, csvLines);

            MessageBox.Show(
                $"Sheet-metal parts found: 1\r\n" +
                $"DWG files saved: {exported}\r\n" +
                $"Failed: {failed}\r\n" +
                $"Folder:\r\n{jobFolder}",
                "DWG Export");
        }

        // ------------------------------------------------------------------
        //  Assembly
        // ------------------------------------------------------------------

        private void RunForAssembly(ISldWorks swApp, IAssemblyDoc asm, string jobFolder)
        {
            if (asm == null)
            {
                MessageBox.Show("The active document is not a valid assembly.", "DWG Export");
                return;
            }

            // Try to resolve lightweight components
            try
            {
                asm.ResolveAllLightWeightComponents(true);
            }
            catch { }

            var usage = CollectSheetMetalUsage(asm);
            if (usage.Count == 0)
            {
                MessageBox.Show("No sheet-metal parts were found in this assembly.", "DWG Export");
                return;
            }

            var csvLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity"
            };

            int exported = 0;
            int failed = 0;

            foreach (var kvp in usage)
            {
                string partPath = kvp.Key;
                SheetMetalUsageInfo info = kvp.Value;

                try
                {
                    // Activate the part document – ExportToDWG2 works reliably only on active docs
                    int actErr = 0;
                    swApp.ActivateDoc3(
                        partPath,
                        false,
                        (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                        ref actErr);

                    var partModel = swApp.IActiveDoc2 as IModelDoc2;
                    if (partModel == null || partModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                    {
                        failed++;
                        continue;
                    }

                    if (!ExportSinglePartToDwg(partModel, partPath, jobFolder, out string dwgFileName))
                    {
                        failed++;
                        continue;
                    }

                    exported++;

                    double thicknessMm = info.ThicknessMeters * 1000.0;
                    csvLines.Add(
                        $"{dwgFileName},{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)},{info.Quantity}");
                    // NEW: close the sheet-metal part we just used, without saving it
                    try
                    {
                        swApp.CloseDoc(partPath); // closes the part document, no save
                    }
                    catch
                    {
                        // ignore any close errors; the assembly stays open
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("DWG export failed for " + partPath + ": " + ex);
                    failed++;
                }
            }

            string csvPath = Path.Combine(jobFolder, "parts.csv");
            TryWriteCsv(csvPath, csvLines);

            MessageBox.Show(
                $"Unique sheet-metal parts found: {usage.Count}\r\n" +
                $"DWG files saved: {exported}\r\n" +
                $"Failed: {failed}\r\n" +
                $"Folder:\r\n{jobFolder}",
                "DWG Export");
        }

        // ------------------------------------------------------------------
        //  Actual DWG export (no post-processing)
        // ------------------------------------------------------------------

        /// <summary>
        /// Exports one sheet-metal part to DWG using ExportToDWG2.
        /// DWG is geometry-only (no bend lines, no sketches).
        /// </summary>
        private static bool ExportSinglePartToDwg(
            IModelDoc2 partModel,
            string modelPath,
            string folder,
            out string dwgFileName)
        {
            dwgFileName = null;

            if (partModel == null || string.IsNullOrEmpty(modelPath) || string.IsNullOrEmpty(folder))
                return false;

            string baseName = Path.GetFileNameWithoutExtension(modelPath);
            if (string.IsNullOrEmpty(baseName))
                baseName = "SheetMetal";

            dwgFileName = baseName + ".dwg";
            string outFile = Path.Combine(folder, dwgFileName);

            var partDoc = partModel as IPartDoc;
            if (partDoc == null)
                return false;

            try
            {
                const int action = (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal;

                // Geometry only – no bend lines, no sketches (laser-cut profile)
                const int sheetMetalOptions = 1;

                bool ok = partDoc.ExportToDWG2(
                    outFile,
                    modelPath,
                    action,
                    true,       // single file
                    null,       // alignment
                    false,      // flip X
                    false,      // flip Y
                    sheetMetalOptions,
                    null);      // views

                return ok;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExportToDWG2 failed for {modelPath}: {ex}");
                return false;
            }
        }

        // ------------------------------------------------------------------
        //  Sheet-metal detection / usage
        // ------------------------------------------------------------------

        private static bool TryGetSheetMetalThickness(IModelDoc2 model, out double thicknessMeters)
        {
            thicknessMeters = GetSheetMetalThicknessMeters(model);
            return thicknessMeters > 0.0;
        }

        private static double GetSheetMetalThicknessMeters(IModelDoc2 model)
        {
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART)
                return 0.0;

            Feature feat = model.FirstFeature() as Feature;
            while (feat != null)
            {
                string typeName = "";
                try { typeName = feat.GetTypeName2(); } catch { }

                if (!string.IsNullOrEmpty(typeName) &&
                    typeName.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        var data = feat.GetDefinition() as SheetMetalFeatureData;
                        if (data != null)
                            return data.Thickness; // meters
                    }
                    catch { }
                }

                object nextObj = null;
                try { nextObj = feat.GetNextFeature(); } catch { }
                feat = nextObj as Feature;
            }

            return 0.0;
        }

        private static bool AssemblyContainsSheetMetal(IAssemblyDoc asm)
        {
            if (asm == null) return false;

            try
            {
                // Make sure lightweight components are resolved so GetModelDoc2 works
                asm.ResolveAllLightWeightComponents(true);
            }
            catch { }

            var model = asm as IModelDoc2;
            if (model == null) return false;

            Configuration conf = null;
            try { conf = (Configuration)model.GetActiveConfiguration(); } catch { }
            if (conf == null) return false;

            IComponent2 root = null;
            try { root = (IComponent2)conf.GetRootComponent3(true); } catch { }
            if (root == null) return false;

            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return TraverseContainsSheetMetal(root, visited);
        }

        private static bool TraverseContainsSheetMetal(IComponent2 comp, HashSet<string> visited)
        {
            if (comp == null || comp.IsSuppressed())
                return false;

            IModelDoc2 model = null;
            try { model = comp.GetModelDoc2() as IModelDoc2; } catch { }

            if (model != null && model.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                string path = null;
                try { path = model.GetPathName(); } catch { }

                if (!string.IsNullOrEmpty(path))
                {
                    if (!visited.Add(path))
                    {
                        // already processed this file; continue to other components
                    }
                }

                if (TryGetSheetMetalThickness(model, out _))
                    return true;
            }

            object childrenObj = null;
            try { childrenObj = comp.GetChildren(); } catch { }
            var children = childrenObj as object[];
            if (children == null) return false;

            foreach (object childObj in children)
            {
                var child = childObj as IComponent2;
                if (child == null) continue;

                if (TraverseContainsSheetMetal(child, visited))
                    return true;
            }

            return false;
        }

        private sealed class SheetMetalUsageInfo
        {
            public double ThicknessMeters;
            public int Quantity;
        }

        /// <summary>
        /// Builds: part path -> usage info (thickness, quantity).
        /// Only sheet-metal parts are included.
        /// </summary>
        private static Dictionary<string, SheetMetalUsageInfo> CollectSheetMetalUsage(IAssemblyDoc asm)
        {
            var result = new Dictionary<string, SheetMetalUsageInfo>(StringComparer.OrdinalIgnoreCase);
            if (asm == null) return result;

            try
            {
                asm.ResolveAllLightWeightComponents(true);
            }
            catch { }

            var model = asm as IModelDoc2;
            if (model == null) return result;

            Configuration conf = null;
            try { conf = (Configuration)model.GetActiveConfiguration(); } catch { }
            if (conf == null) return result;

            IComponent2 root = null;
            try { root = (IComponent2)conf.GetRootComponent3(true); } catch { }
            if (root == null) return result;

            TraverseUsage(root, result);
            return result;
        }

        private static void TraverseUsage(IComponent2 comp, Dictionary<string, SheetMetalUsageInfo> result)
        {
            if (comp == null || comp.IsSuppressed())
                return;

            IModelDoc2 model = null;
            try { model = comp.GetModelDoc2() as IModelDoc2; } catch { }

            if (model != null && model.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                string path = null;
                try { path = model.GetPathName(); } catch { }

                if (!string.IsNullOrWhiteSpace(path) &&
                    TryGetSheetMetalThickness(model, out double thicknessMeters))
                {
                    if (!result.TryGetValue(path, out var info))
                    {
                        info = new SheetMetalUsageInfo
                        {
                            ThicknessMeters = thicknessMeters,
                            Quantity = 0
                        };
                        result.Add(path, info);
                    }

                    info.Quantity++;
                }
            }

            object childrenObj = null;
            try { childrenObj = comp.GetChildren(); } catch { }
            var children = childrenObj as object[];
            if (children == null) return;

            foreach (object childObj in children)
            {
                var child = childObj as IComponent2;
                if (child == null) continue;
                TraverseUsage(child, result);
            }
        }

        // ------------------------------------------------------------------
        //  Folder selection + CSV + active-doc helpers
        // ------------------------------------------------------------------

        private static string SelectMainFolder(IModelDoc2 model)
        {
            string initialDir = null;
            try
            {
                string path = model.GetPathName();
                if (!string.IsNullOrEmpty(path))
                    initialDir = Path.GetDirectoryName(path);
            }
            catch
            {
                // ignore, fall back to default
            }

            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select the MAIN folder for DWG export.";
                dlg.ShowNewFolderButton = true;

                if (!string.IsNullOrEmpty(initialDir))
                    dlg.SelectedPath = initialDir;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.SelectedPath;
            }
        }

        private static string GetDocumentBaseName(IModelDoc2 model)
        {
            if (model == null) return "DWGExport";

            try
            {
                string path = model.GetPathName();
                if (!string.IsNullOrEmpty(path))
                    return Path.GetFileNameWithoutExtension(path);
            }
            catch { }

            try
            {
                string title = model.GetTitle();
                if (!string.IsNullOrEmpty(title))
                {
                    int dot = title.LastIndexOf('.');
                    return dot > 0 ? title.Substring(0, dot) : title;
                }
            }
            catch { }

            return "DWGExport";
        }

        private static void TryWriteCsv(string csvPath, List<string> lines)
        {
            try
            {
                File.WriteAllLines(csvPath, lines);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CSV write failed for " + csvPath + ": " + ex);
            }
        }

        private static string GetActiveDocKey(ISldWorks swApp)
        {
            try
            {
                var doc = swApp.IActiveDoc2 as IModelDoc2;
                if (doc == null) return null;

                string path = doc.GetPathName();
                if (!string.IsNullOrEmpty(path)) return path;

                return doc.GetTitle();
            }
            catch
            {
                return null;
            }
        }

        private static void RestoreActiveDoc(ISldWorks swApp, string key)
        {
            if (string.IsNullOrEmpty(key)) return;

            try
            {
                int err = 0;
                swApp.ActivateDoc3(
                    key,
                    false,
                    (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                    ref err);
            }
            catch { }
        }
    }
}
