using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin
{
    internal class SheetMetalPartInfo
    {
        public string Path { get; }
        public double ThicknessMm { get; }
        public int Quantity { get; set; }

        public SheetMetalPartInfo(string path, double thicknessMm)
        {
            Path = path;
            ThicknessMm = thicknessMm;
            Quantity = 0;
        }
    }

    internal static class SheetMetalAssemblyHelper
    {
        public static Dictionary<string, SheetMetalPartInfo> CollectSheetMetalParts(AssemblyDoc asm)
        {
            var result = new Dictionary<string, SheetMetalPartInfo>(
                StringComparer.OrdinalIgnoreCase);

            var model = (ModelDoc2)asm;
            var conf = (Configuration)model.GetActiveConfiguration();
            var root = (Component2)conf.GetRootComponent3(true);

            TraverseComponent(root, result);
            return result;
        }

        private static void TraverseComponent(
            Component2 comp,
            IDictionary<string, SheetMetalPartInfo> result)
        {
            if (comp == null)
                return;

            if (comp.IsSuppressed())
                return;

            if (comp.IsHidden(true))
                return;

            var childModel = (ModelDoc2)comp.GetModelDoc2();

            if (childModel != null)
            {
                var type = (swDocumentTypes_e)childModel.GetType();

                if (type == swDocumentTypes_e.swDocPART)
                {
                    var part = (PartDoc)childModel;

                    if (TryGetSheetMetalThickness(part, out double thicknessMm))
                    {
                        string path = childModel.GetPathName();
                        if (string.IsNullOrEmpty(path))
                            path = childModel.GetTitle();

                        if (!result.TryGetValue(path, out var info))
                        {
                            info = new SheetMetalPartInfo(path, thicknessMm);
                            result.Add(path, info);
                        }

                        info.Quantity += 1;
                    }
                }
                else if (type == swDocumentTypes_e.swDocASSEMBLY)
                {
                    // Rare, but if you have sub-assemblies loaded as components
                    var subAsm = (AssemblyDoc)childModel;
                    var subConf = (Configuration)childModel.GetActiveConfiguration();
                    var subRoot = (Component2)subConf.GetRootComponent3(true);
                    TraverseComponent(subRoot, result);
                }
            }

            // Recurse children
            object[] children = (object[])comp.GetChildren();
            if (children == null)
                return;

            foreach (Component2 child in children)
                TraverseComponent(child, result);
        }

        private static bool TryGetSheetMetalThickness(PartDoc part, out double thicknessMm)
        {
            thicknessMm = 0.0;

            var model = (ModelDoc2)part;
            Feature feat = (Feature)model.FirstFeature();

            while (feat != null)
            {
                string typeName = feat.GetTypeName2();

                if (string.Equals(typeName, "SheetMetal",
                    StringComparison.OrdinalIgnoreCase))
                {
                    var data = (SheetMetalFeatureData)feat.GetDefinition();
                    double thicknessMeters = data.Thickness;
                    thicknessMm = thicknessMeters * 1000.0;
                    return true;
                }

                feat = (Feature)feat.GetNextFeature();
            }

            return false;
        }
    }
}
