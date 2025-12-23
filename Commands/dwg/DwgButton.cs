using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

using SW2026RibbonAddin.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class DwgButton : IMehdiRibbonButton
    {
        private const int DefaultAssemblyQuantity = 1;
        private int _assemblyQuantity = DefaultAssemblyQuantity;

        public string Id => "DWG";

        public string DisplayName => "DWG";
        public string Tooltip => "Export DWG files for sheet-metal parts";
        public string Hint => "DWG export for sheet metal";

        public string SmallIconFile => "dwg_20.png";
        public string LargeIconFile => "dwg_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 1;

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

            int docType = model.GetType(); // SolidWorks doc type (int)
            bool isPart = docType == (int)swDocumentTypes_e.swDocPART;
            bool isAsm = docType == (int)swDocumentTypes_e.swDocASSEMBLY;

            if (!isPart && !isAsm)
            {
                MessageBox.Show("DWG export is only available for parts and assemblies.", "DWG Export");
                return;
            }

            using (var dlg = new AssemblyQuantityForm(_assemblyQuantity))
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                _assemblyQuantity = dlg.AssemblyQuantity;
            }

            string mainFolder = SelectMainFolder(model);
            if (string.IsNullOrEmpty(mainFolder))
                return;

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

            string originalKey = GetActiveDocKey(swApp);

            try
            {
                if (isPart)
                    RunForSinglePart(swApp, model, jobFolder);
                else
                    RunForAssembly(swApp, (IAssemblyDoc)model, jobFolder);
            }
            finally
            {
                RestoreActiveDoc(swApp, originalKey);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            try
            {
                return context.ActiveModel != null ? AddinContext.Enable : AddinContext.Disable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }

        // ------------------------------------------------------------------
        // Single Part
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
                MessageBox.Show("Save the part before exporting to DWG.", "DWG Export");
                return;
            }

            string cfgName = GetBestConfigName(model, null);
            string material = GetMaterialName(model, cfgName);

            var csvLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity,Material"
            };

            var dwgFileNames = new List<string>();
            int failures = ExportFlatPatternsForPart(
                model,
                modelPath,
                jobFolder,
                uniquePartToken: null,
                globalUsedNames: null,
                dwgFileNames: dwgFileNames);

            int exported = dwgFileNames.Count;
            int totalBodies = exported + failures;
            int failed = failures;

            double thicknessMm = thicknessMeters * 1000.0;

            foreach (string dwgName in dwgFileNames)
            {
                csvLines.Add(
                    $"{CsvCell(dwgName)}," +
                    $"{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)}," +
                    $"{_assemblyQuantity}," +
                    $"{CsvCell(material)}");
            }

            string csvPath = Path.Combine(jobFolder, "parts.csv");
            TryWriteCsv(csvPath, csvLines);

            MessageBox.Show(
                $"Sheet-metal bodies found: {totalBodies}\r\n" +
                $"DWG files saved: {exported}\r\n" +
                $"Failed: {failed}\r\n" +
                $"Material detected: {material}\r\n" +
                $"Folder:\r\n{jobFolder}",
                "DWG Export");
        }

        // ------------------------------------------------------------------
        // Assembly
        // ------------------------------------------------------------------

        private void RunForAssembly(ISldWorks swApp, IAssemblyDoc asm, string jobFolder)
        {
            if (asm == null)
            {
                MessageBox.Show("The active document is not a valid assembly.", "DWG Export");
                return;
            }

            try { asm.ResolveAllLightWeightComponents(true); } catch { }

            var usage = CollectSheetMetalUsage(asm);
            if (usage.Count == 0)
            {
                MessageBox.Show("No sheet-metal parts were found in this assembly.", "DWG Export");
                return;
            }

            var csvLines = new List<string>
            {
                "FileName,PlateThickness_mm,Quantity,Material"
            };

            int exported = 0;
            int failed = 0;
            int totalBodies = 0;

            // Prevent silent overwrites when different parts share the same base filename.
            var globalUsedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var kvp in usage)
            {
                var info = kvp.Value;
                string partPath = info.PartPath;

                try
                {
                    int actErr = 0;

                    // SOLIDWORKS interop here is 4-arg ActivateDoc3 (NOT 5)
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

                    if (!string.IsNullOrWhiteSpace(info.ConfigName))
                    {
                        try { partModel.ShowConfiguration2(info.ConfigName); } catch { }
                    }

                    string material = info.Material;
                    if (string.IsNullOrWhiteSpace(material) || material.Equals("UNKNOWN", StringComparison.OrdinalIgnoreCase))
                        material = GetMaterialName(partModel, info.ConfigName);

                    var dwgFileNames = new List<string>();
                    string uniqueToken = ComputeShortHash(partPath + "||" + (info.ConfigName ?? ""), hexChars: 8);

                    int failuresForPart = ExportFlatPatternsForPart(
                        partModel,
                        partPath,
                        jobFolder,
                        uniquePartToken: uniqueToken,
                        globalUsedNames: globalUsedNames,
                        dwgFileNames: dwgFileNames);

                    int exportedForPart = dwgFileNames.Count;
                    int totalForPart = exportedForPart + failuresForPart;

                    totalBodies += totalForPart;
                    exported += exportedForPart;
                    failed += failuresForPart;

                    double thicknessMm = info.ThicknessMeters * 1000.0;

                    foreach (string dwgName in dwgFileNames)
                    {
                        csvLines.Add(
                            $"{CsvCell(dwgName)}," +
                            $"{thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)}," +
                            $"{info.Quantity * _assemblyQuantity}," +
                            $"{CsvCell(material)}");
                    }

                    try { swApp.CloseDoc(partPath); } catch { }
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
        // DWG Export (flat patterns)
        // ------------------------------------------------------------------

        private static int ExportFlatPatternsForPart(
            IModelDoc2 partModel,
            string modelPath,
            string folder,
            string uniquePartToken,
            HashSet<string> globalUsedNames,
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

            if (flatPatterns.Count == 0)
            {
                string baseName = Path.GetFileNameWithoutExtension(modelPath);
                if (string.IsNullOrEmpty(baseName))
                    baseName = "SheetMetal";

                string stem = MakeExportStem(baseName, uniquePartToken);
                string dwgName = ReserveUniqueFileName(stem + ".dwg", globalUsedNames);
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

            string stemPrefix = MakeExportStem(partBaseName, uniquePartToken);

            foreach (Feature flatFeat in flatPatterns)
            {
                if (flatFeat == null)
                    continue;

                string suffix = null;
                try { suffix = flatFeat.Name; } catch { }

                if (string.IsNullOrWhiteSpace(suffix))
                    suffix = idx.ToString(CultureInfo.InvariantCulture);

                suffix = MakeSafeFilePart(suffix);

                string candidate = $"{stemPrefix}_{suffix}.dwg";
                string finalName = ReserveUniqueFileName(candidate, globalUsedNames, usedNames);

                string outPath = Path.Combine(folder, finalName);

                bool ok = ExportFlatPatternSelected(partDoc, modelPath, flatFeat, outPath);
                if (ok)
                    dwgFileNames.Add(finalName);
                else
                    failures++;

                idx++;
            }

            return failures;
        }

        private static string MakeExportStem(string baseName, string uniquePartToken)
        {
            baseName = (baseName ?? "").Trim();
            if (baseName.Length == 0)
                baseName = "SheetMetal";

            uniquePartToken = (uniquePartToken ?? "").Trim();
            if (uniquePartToken.Length == 0)
                return baseName;

            return baseName + "__" + uniquePartToken;
        }

        private static string ReserveUniqueFileName(string candidate, HashSet<string> globalUsedNames, HashSet<string> localUsedNames = null)
        {
            candidate = (candidate ?? "").Trim();
            if (candidate.Length == 0)
                candidate = "SheetMetal.dwg";

            string ext = Path.GetExtension(candidate);
            if (string.IsNullOrEmpty(ext))
                ext = ".dwg";

            string stem = Path.GetFileNameWithoutExtension(candidate);
            if (string.IsNullOrEmpty(stem))
                stem = "SheetMetal";

            string finalName = stem + ext;
            int nIdx = 1;

            while ((localUsedNames != null && localUsedNames.Contains(finalName)) ||
                   (globalUsedNames != null && globalUsedNames.Contains(finalName)))
            {
                finalName = stem + "_" + nIdx.ToString(CultureInfo.InvariantCulture) + ext;
                nIdx++;
            }

            if (localUsedNames != null)
                localUsedNames.Add(finalName);
            if (globalUsedNames != null)
                globalUsedNames.Add(finalName);

            return finalName;
        }

        private static string ComputeShortHash(string input, int hexChars)
        {
            input = input ?? "";
            if (hexChars <= 0)
                hexChars = 8;

            try
            {
                using (var sha = SHA256.Create())
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(input);
                    byte[] hash = sha.ComputeHash(bytes);

                    var sb = new StringBuilder(hash.Length * 2);
                    foreach (byte b in hash)
                    {
                        sb.Append(b.ToString("X2"));
                        if (sb.Length >= hexChars)
                            break;
                    }

                    if (sb.Length > hexChars)
                        sb.Length = hexChars;

                    return sb.ToString();
                }
            }
            catch
            {
                // Worst-case fallback: stable-ish sanitization of the input.
                return Math.Abs(input.GetHashCode()).ToString("X");
            }
        }

        private static bool ExportFlatPatternWithoutSelection(IPartDoc partDoc, string modelPath, string outFile)
        {
            try
            {
                const int action = (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal;
                const int sheetMetalOptions = 1; // geometry only

                bool ok = partDoc.ExportToDWG2(
                    outFile,
                    modelPath,
                    action,
                    true,
                    null,
                    false,
                    false,
                    sheetMetalOptions,
                    null);

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
                    true,
                    null,
                    false,
                    false,
                    sheetMetalOptions,
                    null);

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
            try { feat = model.FirstFeature() as Feature; } catch { }

            while (feat != null)
            {
                CollectFlatPatternsRecursive(feat, result);

                Feature next = null;
                try { next = feat.GetNextFeature() as Feature; } catch { }
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
                    result.Add(feat);
            }
            catch { }

            Feature sub = null;
            try { sub = feat.GetFirstSubFeature() as Feature; } catch { }

            while (sub != null)
            {
                CollectFlatPatternsRecursive(sub, result);

                Feature next = null;
                try { next = sub.GetNextSubFeature() as Feature; } catch { }
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
        // Sheet-metal detection / thickness
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
                            return data.Thickness;
                    }
                    catch { }
                }

                Feature next = null;
                try { next = feat.GetNextFeature() as Feature; } catch { }
                feat = next;
            }

            return 0.0;
        }

        private sealed class SheetMetalUsageInfo
        {
            public string PartPath;
            public string ConfigName;
            public string Material;
            public double ThicknessMeters;
            public int Quantity;
        }

        private static Dictionary<string, SheetMetalUsageInfo> CollectSheetMetalUsage(IAssemblyDoc asm)
        {
            var result = new Dictionary<string, SheetMetalUsageInfo>(StringComparer.OrdinalIgnoreCase);
            if (asm == null) return result;

            object compsObj = null;
            try { compsObj = asm.GetComponents(false); } catch { }

            var comps = compsObj as object[];
            if (comps == null) return result;

            foreach (object o in comps)
            {
                var comp = o as IComponent2;
                if (comp == null || comp.IsSuppressed()) continue;

                IModelDoc2 refModel = null;
                try { refModel = comp.GetModelDoc2() as IModelDoc2; } catch { }
                if (refModel == null || refModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                    continue;

                string path = null;
                try { path = refModel.GetPathName(); } catch { }
                if (string.IsNullOrWhiteSpace(path))
                    continue;

                if (!TryGetSheetMetalThickness(refModel, out double thicknessMeters))
                    continue;

                string cfgName = GetBestConfigName(refModel, comp);
                string dictKey = (path ?? "").Trim().ToUpperInvariant() + "||" + (cfgName ?? "").Trim().ToUpperInvariant();

                if (!result.TryGetValue(dictKey, out var info))
                {
                    string material = GetMaterialName(refModel, cfgName);

                    info = new SheetMetalUsageInfo
                    {
                        PartPath = path,
                        ConfigName = cfgName,
                        Material = material,
                        ThicknessMeters = thicknessMeters,
                        Quantity = 0
                    };
                    result[dictKey] = info;
                }

                info.Quantity++;
            }

            return result;
        }

        // ------------------------------------------------------------------
        // Material detection (robust)
        // ------------------------------------------------------------------

        private static string GetBestConfigName(IModelDoc2 model, IComponent2 comp)
        {
            if (comp != null)
            {
                try
                {
                    string rc = comp.ReferencedConfiguration;
                    if (!string.IsNullOrWhiteSpace(rc))
                        return rc;
                }
                catch { }
            }

            try
            {
                var cfg = model?.ConfigurationManager?.ActiveConfiguration;
                if (cfg != null && !string.IsNullOrWhiteSpace(cfg.Name))
                    return cfg.Name;
            }
            catch { }

            return "";
        }

        private static string GetMaterialName(IModelDoc2 model, string configName)
        {
            string mat =
                TryGetMaterial_FromExtension(model, configName) ??
                TryGetMaterial_FromMaterialIdName(model) ??
                TryGetMaterial_FromCustomProps(model, configName) ??
                TryGetMaterial_FromFirstSolidBody(model, configName);

            if (string.IsNullOrWhiteSpace(mat))
                return "UNKNOWN";

            mat = mat.Trim();
            if (mat.Equals("NOT SPECIFIED", StringComparison.OrdinalIgnoreCase))
                return "UNKNOWN";

            return mat;
        }

        private static string TryGetMaterial_FromExtension(IModelDoc2 model, string configName)
        {
            try
            {
                var ext = model?.Extension;
                if (ext == null) return null;

                // Reflection on interface type is safe
                var mi = typeof(IModelDocExtension).GetMethod("GetMaterialPropertyName2");
                if (mi == null) return null;

                var ps = mi.GetParameters();

                // Variant A: string GetMaterialPropertyName2(string configName, out string dbName)
                if (ps.Length == 2 && ps[0].ParameterType == typeof(string) && ps[1].IsOut)
                {
                    object[] args = new object[] { configName ?? "", null };
                    var ret = mi.Invoke(ext, args);
                    var mat = ret as string;
                    if (!string.IsNullOrWhiteSpace(mat))
                        return mat;
                }

                // Variant B: bool GetMaterialPropertyName2(string configName, out string dbName, out string matName)
                if (ps.Length == 3 && ps[0].ParameterType == typeof(string) && ps[1].IsOut && ps[2].IsOut)
                {
                    object[] args = new object[] { configName ?? "", null, null };
                    mi.Invoke(ext, args);

                    var mat = args[2] as string;
                    if (!string.IsNullOrWhiteSpace(mat))
                        return mat;

                    var maybe = args[1] as string;
                    if (!string.IsNullOrWhiteSpace(maybe))
                        return maybe;
                }
            }
            catch { }

            return null;
        }

        private static string TryGetMaterial_FromMaterialIdName(IModelDoc2 model)
        {
            try
            {
                // ✅ FIX #1: NEVER use model.GetType() here (SW GetType() returns int).
                // Reflect on the interface instead.
                var p = typeof(IModelDoc2).GetProperty("MaterialIdName");
                if (p == null)
                    return null;

                var id = p.GetValue(model, null) as string;
                if (string.IsNullOrWhiteSpace(id))
                    return null;

                // Often: "SOLIDWORKS Materials.sldmat::AISI 304"
                int idx = id.LastIndexOf("::", StringComparison.Ordinal);
                string mat = (idx >= 0) ? id.Substring(idx + 2) : id;
                mat = (mat ?? "").Trim();
                return string.IsNullOrWhiteSpace(mat) ? null : mat;
            }
            catch { }

            return null;
        }

        private static string TryGetMaterial_FromCustomProps(IModelDoc2 model, string configName)
        {
            try
            {
                var ext = model?.Extension;
                if (ext == null) return null;

                string v = TryGetCustomProp(ext, configName, "SW-Material")
                        ?? TryGetCustomProp(ext, configName, "Material");

                if (!string.IsNullOrWhiteSpace(v))
                    return v;

                v = TryGetCustomProp(ext, "", "SW-Material")
                 ?? TryGetCustomProp(ext, "", "Material");

                return string.IsNullOrWhiteSpace(v) ? null : v;
            }
            catch { }

            return null;
        }

        private static string TryGetCustomProp(IModelDocExtension ext, string configName, string propName)
        {
            try
            {
                if (ext == null) return null;

                CustomPropertyManager cpm = null;
                try { cpm = ext.CustomPropertyManager[configName ?? ""]; } catch { cpm = null; }
                if (cpm == null) return null;

                string valOut, resolved;

                // ✅ FIX #2: don't assume return type (some interop return bool, some int)
                cpm.Get4(propName, false, out valOut, out resolved);

                string v = !string.IsNullOrWhiteSpace(resolved) ? resolved : valOut;
                v = (v ?? "").Trim();
                return string.IsNullOrWhiteSpace(v) ? null : v;
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetMaterial_FromFirstSolidBody(IModelDoc2 model, string configName)
        {
            try
            {
                var part = model as IPartDoc;
                if (part == null) return null;

                object bodiesObj = null;
                try { bodiesObj = part.GetBodies2((int)swBodyType_e.swSolidBody, false); } catch { bodiesObj = null; }

                var bodies = bodiesObj as object[];
                if (bodies == null || bodies.Length == 0)
                    return null;

                var body = bodies[0];
                if (body == null)
                    return null;

                // Use reflection (body is object, safe)
                var mi = body.GetType().GetMethod("GetMaterialPropertyName2");
                if (mi == null) return null;

                var ps = mi.GetParameters();

                // Variant B: bool GetMaterialPropertyName2(string config, out string db, out string mat)
                if (ps.Length == 3 && ps[0].ParameterType == typeof(string) && ps[1].IsOut && ps[2].IsOut)
                {
                    object[] args = new object[] { configName ?? "", null, null };
                    mi.Invoke(body, args);
                    var mat = args[2] as string;
                    if (!string.IsNullOrWhiteSpace(mat))
                        return mat;
                }

                // Variant A: string GetMaterialPropertyName2(string config, out string db)
                if (ps.Length == 2 && ps[0].ParameterType == typeof(string) && ps[1].IsOut)
                {
                    object[] args = new object[] { configName ?? "", null };
                    var ret = mi.Invoke(body, args);
                    var mat = ret as string;
                    if (!string.IsNullOrWhiteSpace(mat))
                        return mat;
                }
            }
            catch { }

            return null;
        }

        // ------------------------------------------------------------------
        // Folder / CSV / doc helpers
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
            catch { }

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
                File.WriteAllLines(csvPath, lines, Encoding.UTF8);
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

        private static string CsvCell(string s)
        {
            if (s == null) return "";
            s = s.Trim();

            bool needsQuotes =
                s.Contains(",") ||
                s.Contains("\"") ||
                s.Contains("\r") ||
                s.Contains("\n");

            if (!needsQuotes)
                return s;

            s = s.Replace("\"", "\"\"");
            return "\"" + s + "\"";
        }
    }
}
