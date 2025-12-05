using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SW2026RibbonAddin.Forms;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// DWG export button:
    /// - For a sheet-metal part (single-body or multi-body): exports one DWG per flat pattern
    ///   (per sheet-metal body) into a new folder.
    /// - For an assembly: exports one DWG per flat pattern for each unique sheet-metal part
    ///   (recursively through sub-assemblies).
    /// - Uses FolderBrowserDialog to pick a MAIN folder, then creates a subfolder named
    ///   after the active document (must NOT already exist).
    /// - Writes a CSV inside that subfolder with columns:
    ///   FileName,PlateThickness_mm,Quantity.
    /// </summary>
    internal sealed class DwgButton : IMehdiRibbonButton
    {
        // Default number of assemblies if user just clicks OK without changes
        private const int DefaultAssemblyQuantity = 1;

        // Filled from the dialog every time the DWG command runs
        private int _assemblyQuantity = DefaultAssemblyQuantity;

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


            // Ask for number of assemblies (used to scale CSV quantities)
            using (var dlg = new AssemblyQuantityForm(_assemblyQuantity))
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return; // user cancelled

                _assemblyQuantity = dlg.AssemblyQuantity;
            }

            // Let user pick the MAIN folder (standard folder dialog)
            string mainFolder = SelectMainFolder(model);
            if (string.IsNullOrEmpty(mainFolder))
                return; // user cancelled

            // Create subfolder named after the active document
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
            // New strategy: keep DWG button enabled whenever some document is active.
            try
            {
                return context.ActiveModel != null
                    ? AddinContext.Enable
                    : AddinContext.Disable;
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

            var dwgFileNames = new List<string>();
            int failures = ExportFlatPatternsForPart(model, modelPath, jobFolder, dwgFileNames);

            int exported = dwgFileNames.Count;
            int totalBodies = exported + failures;
            int failed = failures;

            double thicknessMm = thicknessMeters * 1000.0;

            foreach (string dwgName in dwgFileNames)
            {
                csvLines.Add(
                    $"{dwgName},{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)},{_assemblyQuantity}");
            }

            string csvPath = Path.Combine(jobFolder, "parts.csv");
            TryWriteCsv(csvPath, csvLines);

            MessageBox.Show(
                $"Sheet-metal bodies found: {totalBodies}\r\n" +
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
            int totalBodies = 0;

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

                    var dwgFileNames = new List<string>();
                    int failuresForPart = ExportFlatPatternsForPart(partModel, partPath, jobFolder, dwgFileNames);

                    int exportedForPart = dwgFileNames.Count;
                    int totalForPart = exportedForPart + failuresForPart;

                    totalBodies += totalForPart;
                    exported += exportedForPart;
                    failed += failuresForPart;

                    double thicknessMm = info.ThicknessMeters * 1000.0;

                    foreach (string dwgName in dwgFileNames)
                    {
                        csvLines.Add(
                            $"{dwgName},{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)},{info.Quantity*_assemblyQuantity}");
                    }

                    // Close the part we opened for export (without saving)
                    try
                    {
                        swApp.CloseDoc(partPath);
                    }
                    catch
                    {
                        // ignore close errors
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
                $"Total sheet-metal bodies (plates): {totalBodies}\r\n" +
                $"DWG files saved: {exported}\r\n" +
                $"Failed: {failed}\r\n" +
                $"Folder:\r\n{jobFolder}",
                "DWG Export");
        }

        // ------------------------------------------------------------------
        //  Actual DWG export (handles multi-body via flat patterns)
        // ------------------------------------------------------------------

        /// <summary>
        /// Exports all flat patterns (sheet-metal bodies) in the given part to DWG files.
        /// For single-body sheet-metal, this produces one DWG (same as before).
        /// For multi-body sheet-metal, this produces one DWG per flat pattern.
        /// Returns the number of failed bodies; successful DWG file names are added
        /// to <paramref name="dwgFileNames"/>.
        /// </summary>
        private static int ExportFlatPatternsForPart(
            IModelDoc2 partModel,
            string modelPath,
            string folder,
            List<string> dwgFileNames)
        {
            if (partModel == null)
                return 1;

            var partDoc = partModel as IPartDoc;
            if (partDoc == null)
                return 1;

            if (string.IsNullOrEmpty(modelPath))
                return 1;

            var flatPatterns = GetFlatPatternFeatures(partModel);

            // If no flat patterns found, fall back to a single export
            if (flatPatterns.Count == 0)
            {
                string baseName = Path.GetFileNameWithoutExtension(modelPath);
                if (string.IsNullOrEmpty(baseName))
                    baseName = "SheetMetal";

                string dwgName = baseName + ".dwg";
                string outPath = Path.Combine(folder, dwgName);

                if (ExportFlatPatternWithoutSelection(partDoc, modelPath, outPath))
                {
                    dwgFileNames.Add(dwgName);
                    return 0;
                }

                return 1;
            }

            string partBaseName = Path.GetFileNameWithoutExtension(modelPath);
            if (string.IsNullOrEmpty(partBaseName))
                partBaseName = "SheetMetal";

            int failures = 0;
            int idx = 1;
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Feature flatFeat in flatPatterns)
            {
                if (flatFeat == null)
                    continue;

                string suffix = null;
                try { suffix = flatFeat.Name; }
                catch { }

                if (string.IsNullOrWhiteSpace(suffix))
                    suffix = idx.ToString(CultureInfo.InvariantCulture);

                suffix = MakeSafeFilePart(suffix);

                string candidate = $"{partBaseName}_{suffix}.dwg";
                string finalName = candidate;
                int n = 1;
                while (!usedNames.Add(finalName))
                {
                    finalName = $"{partBaseName}_{suffix}_{n}.dwg";
                    n++;
                }

                string outPath = Path.Combine(folder, finalName);

                bool ok = ExportFlatPatternSelected(partDoc, modelPath, flatFeat, outPath);
                if (ok)
                {
                    dwgFileNames.Add(finalName);
                }
                else
                {
                    failures++;
                }

                idx++;
            }

            return failures;
        }

        private static bool ExportFlatPatternWithoutSelection(IPartDoc partDoc, string modelPath, string outFile)
        {
            try
            {
                const int action = (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal;
                const int sheetMetalOptions = 1; // geometry only – no bend lines, no sketches

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
                Debug.WriteLine("ExportFlatPatternWithoutSelection failed: " + ex);
                return false;
            }
        }

        private static bool ExportFlatPatternSelected(IPartDoc partDoc, string modelPath, Feature flatPatternFeat, string outFile)
        {
            try
            {
                if (flatPatternFeat == null)
                    return false;

                bool selOk = flatPatternFeat.Select2(false, -1);
                if (!selOk)
                    return false;

                const int action = (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal;
                const int sheetMetalOptions = 1; // geometry only

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
                Debug.WriteLine("ExportFlatPatternSelected failed: " + ex);
                return false;
            }
        }

        private static List<Feature> GetFlatPatternFeatures(IModelDoc2 model)
        {
            var result = new List<Feature>();
            if (model == null)
                return result;

            Feature feat = null;
            try { feat = model.FirstFeature() as Feature; }
            catch { }

            while (feat != null)
            {
                CollectFlatPatternsRecursive(feat, result);

                Feature next = null;
                try { next = feat.GetNextFeature() as Feature; }
                catch { }
                feat = next;
            }

            return result;
        }

        private static void CollectFlatPatternsRecursive(Feature feat, List<Feature> result)
        {
            if (feat == null)
                return;

            try
            {
                string typeName = feat.GetTypeName2();
                if (string.Equals(typeName, "FlatPattern", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(feat);
                }
            }
            catch { }

            Feature sub = null;
            try { sub = feat.GetFirstSubFeature() as Feature; }
            catch { }

            while (sub != null)
            {
                CollectFlatPatternsRecursive(sub, result);

                Feature next = null;
                try { next = sub.GetNextSubFeature() as Feature; }
                catch { }
                sub = next;
            }
        }

        private static string MakeSafeFilePart(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Body";

            char[] invalid = Path.GetInvalidFileNameChars();
            char[] chars = name.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(invalid, chars[i]) >= 0)
                    chars[i] = '_';
            }

            return new string(chars);
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
                try { typeName = feat.GetTypeName2(); }
                catch { }

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
                try { nextObj = feat.GetNextFeature(); }
                catch { }
                feat = nextObj as Feature;
            }

            return 0.0;
        }

        private sealed class SheetMetalUsageInfo
        {
            public double ThicknessMeters;
            public int Quantity;
        }

        /// <summary>
        /// Builds: part path -> usage info (thickness, quantity of that part in assembly).
        /// Only parts that contain sheet-metal features are included.
        /// </summary>
        private static Dictionary<string, SheetMetalUsageInfo> CollectSheetMetalUsage(IAssemblyDoc asm)
        {
            var result = new Dictionary<string, SheetMetalUsageInfo>(StringComparer.OrdinalIgnoreCase);
            if (asm == null) return result;

            object compsObj = null;
            try { compsObj = asm.GetComponents(false); } // all levels
            catch { }

            var comps = compsObj as object[];
            if (comps == null) return result;

            foreach (object o in comps)
            {
                var comp = o as IComponent2;
                if (comp == null || comp.IsSuppressed()) continue;

                IModelDoc2 refModel = null;
                try { refModel = comp.GetModelDoc2() as IModelDoc2; }
                catch { }

                if (refModel == null || refModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                    continue;

                string path = null;
                try { path = refModel.GetPathName(); }
                catch { }

                if (string.IsNullOrWhiteSpace(path))
                    continue;

                if (!TryGetSheetMetalThickness(refModel, out double thicknessMeters))
                    continue;

                if (!result.TryGetValue(path, out var info))
                {
                    info = new SheetMetalUsageInfo
                    {
                        ThicknessMeters = thicknessMeters,
                        Quantity = 0
                    };
                    result[path] = info;
                }

                info.Quantity++;
            }

            return result;
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
