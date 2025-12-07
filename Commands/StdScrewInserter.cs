using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Represents one screw in the STD folder (file path, nominal diameter and length in mm).
    /// </summary>
    internal sealed class StdScrewDefinition
    {
        public StdScrewDefinition(string filePath, double diameterMm, double lengthMm)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            DiameterMm = diameterMm;
            LengthMm = lengthMm;
        }

        public string FilePath { get; }
        public double DiameterMm { get; }
        public double LengthMm { get; }

        public override string ToString()
        {
            return $"{Path.GetFileName(FilePath)} (M{DiameterMm}×{LengthMm})";
        }
    }

    /// <summary>
    /// Loads and caches screws from the STD folder located next to the add-in DLL.
    /// </summary>
    internal static class StdScrewLibrary
    {
        private static List<StdScrewDefinition> _screws;
        private static string _lastBaseDir;

        /// <summary>
        /// Legacy property – used by other methods like FindBestScrew.
        /// It just returns the currently cached list.
        /// </summary>
        public static List<StdScrewDefinition> Screws
        {
            get
            {
                // Avoid null – always return a list
                return _screws ?? (_screws = new List<StdScrewDefinition>());
            }
        }

        /// <summary>
        /// Get list of screws for the given model.
        /// It first tries STD next to the active document,
        /// then falls back to STD next to the add‑in DLL.
        /// </summary>
        public static List<StdScrewDefinition> GetScrewsForModel(IModelDoc2 model)
        {
            string baseDir = GetBestBaseDirectory(model);

            // Reuse cache if base directory did not change
            if (_screws != null &&
                string.Equals(baseDir, _lastBaseDir, StringComparison.OrdinalIgnoreCase))
            {
                return _screws;
            }

            _lastBaseDir = baseDir;
            _screws = LoadFromDisk(baseDir);
            return _screws;
        }

        /// <summary>
        /// Decide where to look for STD:
        ///  1) next to active document, if it has a path and a STD subfolder
        ///  2) next to the add‑in DLL
        /// </summary>
        private static string GetBestBaseDirectory(IModelDoc2 model)
        {
            // 1) Try folder of active document
            if (model != null)
            {
                string docPath = model.GetPathName();
                if (!string.IsNullOrWhiteSpace(docPath))
                {
                    string docDir = Path.GetDirectoryName(docPath);
                    if (!string.IsNullOrWhiteSpace(docDir))
                    {
                        string stdInDocDir = Path.Combine(docDir, "STD");
                        if (Directory.Exists(stdInDocDir))
                            return docDir;
                    }
                }
            }

            // 2) Fallback: folder of the add‑in DLL
            string asmDir = Path.GetDirectoryName(typeof(Addin).Assembly.Location);
            return asmDir ?? System.Environment.CurrentDirectory;
        }

        private static List<StdScrewDefinition> LoadFromDisk(string baseDir)
        {
            var result = new List<StdScrewDefinition>();

            if (string.IsNullOrWhiteSpace(baseDir))
                return result;

            // Look in <baseDir>\STD and all its subfolders
            string stdDir = Path.Combine(baseDir, "STD");
            if (!Directory.Exists(stdDir))
                return result;

            try
            {
                foreach (string file in Directory.GetFiles(stdDir, "*.SLDPRT", SearchOption.AllDirectories))
                {
                    if (TryParseScrewFile(file, out StdScrewDefinition def))
                        result.Add(def);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to load STD library: " + ex);
            }

            return result;
        }

        /// <summary>
        /// Parse filenames like:
        ///   411006-000100-00 (ISO 4017) M6×20-8.8.SLDPRT
        ///   411002-000100-00 (ISO 4017) M5x10-8.8.SLDPRT
        /// </summary>
        private static bool TryParseScrewFile(string filePath, out StdScrewDefinition screw)
        {
            screw = null;

            string name = Path.GetFileNameWithoutExtension(filePath) ?? string.Empty;

            // Matches ...M6×20... or ...M5x10...
            Match m = Regex.Match(name, @"M(?<d>\d+)[x×](?<L>\d+)", RegexOptions.IgnoreCase);
            if (!m.Success)
                return false;

            if (!double.TryParse(m.Groups["d"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double dMm))
                return false;
            if (!double.TryParse(m.Groups["L"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double lMm))
                return false;

            screw = new StdScrewDefinition(filePath, dMm, lMm);
            return true;
        }
    }

    /// <summary>
    /// One cylindrical face of a hole stack.
    /// </summary>
    internal sealed class HoleFaceInfo
    {
        public HoleFaceInfo(IFace2 face, double[] origin, double[] axisUnit, double radius)
        {
            Face = face ?? throw new ArgumentNullException(nameof(face));
            Origin = origin ?? throw new ArgumentNullException(nameof(origin));
            AxisUnit = axisUnit ?? throw new ArgumentNullException(nameof(axisUnit));
            Radius = radius;
        }

        public IFace2 Face { get; }
        public double[] Origin { get; }   // point on axis (meters)
        public double[] AxisUnit { get; } // unit vector along axis
        public double Radius { get; }     // meters
    }

    /// <summary>
    /// Group of coaxial cylindrical faces that form one hole stack (through multiple plates).
    /// </summary>
    internal sealed class HoleStack
    {
        public List<HoleFaceInfo> Faces { get; } = new List<HoleFaceInfo>();

        public bool IsEmpty => Faces.Count == 0;

        public double[] AxisUnit => Faces.Count == 0 ? new[] { 0.0, 0.0, 1.0 } : Faces[0].AxisUnit;

        public double[] AxisOrigin => Faces.Count == 0 ? new[] { 0.0, 0.0, 0.0 } : Faces[0].Origin;

        public double MinRadius
        {
            get
            {
                if (Faces.Count == 0)
                    return 0.0;

                double r = Faces[0].Radius;
                for (int i = 1; i < Faces.Count; i++)
                {
                    if (Faces[i].Radius < r)
                        r = Faces[i].Radius;
                }

                return r;
            }
        }

        public double DiameterMm => MinRadius * 2.0 * 1000.0;

        public IFace2 GetReferenceCylindricalFace()
        {
            if (Faces.Count == 0)
                return null;

            HoleFaceInfo best = Faces[0];

            for (int i = 1; i < Faces.Count; i++)
            {
                if (Faces[i].Radius < best.Radius)
                    best = Faces[i];
            }

            return best.Face;
        }
    }

    /// <summary>
    /// Core logic: read selected cylindrical faces, compute hole diameter & height,
    /// pick suitable STD screw, insert component, and add mates.
    /// </summary>
    internal static class StdScrewInserter
    {
        // Tolerances and allowances
        private const double AxisDistanceTolerance = 1e-4;  // 0.1 mm, in meters
        private const double DiameterSlackMm = 0.6;         // allowed oversize vs screw nominal
        private const double LengthAllowanceMm = 2.0;       // extra length beyond stack height

        /// <summary>
        /// Entry point called by the ribbon button.
        /// </summary>
        public static void Run(AddinContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));

            var swApp = context.SwApp;
            var model = context.ActiveModel;

            if (model == null)
            {
                MessageBox.Show(
                    "No active document.\r\n\r\nOpen an assembly, select the cylindrical faces of the hole(s) and run the command again.",
                    "Insert STD Screw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (!(model is IAssemblyDoc asmDoc))
            {
                MessageBox.Show(
                    "Insert STD Screw only works in assemblies.\r\n\r\nOpen an assembly, select the cylindrical faces of the hole(s) and run the command again.",
                    "Insert STD Screw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var screws = StdScrewLibrary.GetScrewsForModel(model);
            if (screws.Count == 0)
            {
                MessageBox.Show(
                    "No screw parts were found.\r\n\r\n" +
                    "I looked for a folder named \"STD\" next to the active document and next to the add-in DLL.\r\n" +
                    "Make sure your screw parts (*.SLDPRT) are inside one of these locations:\r\n" +
                    "  - <your assembly folder>\\STD\\\r\n" +
                    "  - <add-in DLL folder>\\STD\\",
                    "Insert STD Screw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }


            List<HoleStack> stacks = CollectSelectedHoleStacks(model);
            if (stacks.Count == 0)
            {
                MessageBox.Show(
                    "No suitable cylindrical faces were found in the current selection.\r\n\r\n" +
                    "Usage:\r\n" +
                    "  1. In the assembly, select all cylindrical faces of the hole stack (across all plates).\r\n" +
                    "  2. Run \"Insert STD Screw\".\r\n\r\n" +
                    "You may select multiple concentric stacks at the same time.",
                    "Insert STD Screw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var swModel = (IModelDoc2)asmDoc;
            int insertedCount = 0;

            foreach (HoleStack stack in stacks)
            {
                try
                {
                    if (ProcessHoleStack(swApp, asmDoc, swModel, stack))
                        insertedCount++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Error while inserting screw for stack: " + ex);
                    swApp.SendMsgToUser2(
                        "Failed to insert STD screw for one of the selected holes:\r\n\r\n" + ex.Message,
                        (int)swMessageBoxIcon_e.swMbStop,
                        (int)swMessageBoxBtn_e.swMbOk);
                }
            }

            if (insertedCount > 0)
            {
                swModel.EditRebuild3();
            }
        }

        /// <summary>
        /// Handles one coaxial stack. Returns true if a screw was inserted.
        /// </summary>
        private static bool ProcessHoleStack(
            SldWorks swApp,
            IAssemblyDoc asmDoc,
            IModelDoc2 swModel,
            HoleStack stack)
        {
            if (stack == null || stack.IsEmpty)
                return false;

            ComputeStackGeometry(
                stack,
                out double holeDiameterMm,
                out double stackHeightMm,
                out double[] axisOrigin,
                out double[] axisDirUnit);

            if (holeDiameterMm <= 0.0 || stackHeightMm <= 0.0)
            {
                swApp.SendMsgToUser2(
                    "Could not measure the selected hole. Check that only cylindrical faces are selected.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                return false;
            }

            double requiredLengthMm = stackHeightMm + LengthAllowanceMm;

            StdScrewDefinition screw = FindBestScrew(holeDiameterMm, requiredLengthMm);
            if (screw == null)
            {
                swApp.SendMsgToUser2(
                    $"No suitable screw found in STD folder for hole Ø{holeDiameterMm:F2} mm and stack height {stackHeightMm:F1} mm.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                return false;
            }

            // 1) Make sure the screw model is loaded into memory
            int openErr, openWarn;
            IModelDoc2 screwDoc = EnsureScrewModelLoaded(swApp, screw.FilePath, out openErr, out openWarn);
            if (screwDoc == null)
            {
                swApp.SendMsgToUser2(
                    "Failed to open screw model:\r\n" + screw.FilePath +
                    (openErr != 0 ? $"\r\nError code: {openErr}" : ""),
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                return false;
            }

            // 2) Now add the component to the assembly
            Component2 comp = asmDoc.AddComponent5(
                screw.FilePath,
                (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                "",
                false,
                "",
                axisOrigin[0],
                axisOrigin[1],
                axisOrigin[2]);

            if (comp == null)
            {
                swApp.SendMsgToUser2(
                    "Failed to insert component:\r\n" + screw.FilePath,
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                return false;
            }

            // Concentric mate between smallest cylindrical face of the hole and screw shank
            AddConcentricMate(swModel, asmDoc, comp, stack, screw);

            swApp.SendMsgToUser2(
                $"Inserted {Path.GetFileNameWithoutExtension(screw.FilePath)} for hole Ø{holeDiameterMm:F2}×{stackHeightMm:F1} mm.",
                (int)swMessageBoxIcon_e.swMbInformation,
                (int)swMessageBoxBtn_e.swMbOk);

            return true;
        }

        /// <summary>
        /// Read selected cylindrical faces and group them into coaxial stacks.
        /// </summary>
        private static List<HoleStack> CollectSelectedHoleStacks(IModelDoc2 model)
        {
            var result = new List<HoleStack>();

            SelectionMgr selMgr = model.ISelectionManager;
            if (selMgr == null)
                return result;

            int count = selMgr.GetSelectedObjectCount2(-1);
            if (count <= 0)
                return result;

            var faces = new List<HoleFaceInfo>();

            for (int i = 1; i <= count; i++)
            {
                int selType = selMgr.GetSelectedObjectType3(i, -1);
                if (selType != (int)swSelectType_e.swSelFACES)
                    continue;

                IFace2 face = selMgr.GetSelectedObject6(i, -1) as IFace2;
                if (face == null)
                    continue;

                ISurface surf = face.IGetSurface();
                if (surf == null || !surf.IsCylinder())
                    continue;

                double[] cyl = surf.CylinderParams as double[];
                if (cyl == null || cyl.Length < 7)
                    continue;

                double[] origin = new[] { cyl[0], cyl[1], cyl[2] };
                double[] axis = new[] { cyl[3], cyl[4], cyl[5] };
                double radius = Math.Abs(cyl[6]);

                if (Norm(axis) < 1e-9)
                    continue;

                Normalize(axis);

                faces.Add(new HoleFaceInfo(face, origin, axis, radius));
            }

            // Group by coaxiality
            foreach (HoleFaceInfo fi in faces)
            {
                HoleStack found = null;

                foreach (HoleStack stack in result)
                {
                    if (AreCoaxial(fi, stack.Faces[0]))
                    {
                        found = stack;
                        break;
                    }
                }

                if (found == null)
                {
                    found = new HoleStack();
                    result.Add(found);
                }

                found.Faces.Add(fi);
            }

            return result;
        }

        /// <summary>
        /// Measure diameter and total height of the stack along its axis.
        /// </summary>
        private static void ComputeStackGeometry(
            HoleStack stack,
            out double holeDiameterMm,
            out double stackHeightMm,
            out double[] axisOrigin,
            out double[] axisDirUnit)
        {
            holeDiameterMm = 0.0;
            stackHeightMm = 0.0;
            axisOrigin = new[] { 0.0, 0.0, 0.0 };
            axisDirUnit = new[] { 0.0, 0.0, 1.0 };

            if (stack.IsEmpty)
                return;

            HoleFaceInfo first = stack.Faces[0];
            axisOrigin = (double[])first.Origin.Clone();
            axisDirUnit = (double[])first.AxisUnit.Clone();

            holeDiameterMm = stack.DiameterMm;

            bool any = false;
            double tMin = 0.0;
            double tMax = 0.0;

            foreach (HoleFaceInfo fi in stack.Faces)
            {
                double[] box = fi.Face.GetBox() as double[];
                if (box == null || box.Length < 6)
                    continue;

                double xMin = box[0];
                double yMin = box[1];
                double zMin = box[2];
                double xMax = box[3];
                double yMax = box[4];
                double zMax = box[5];

                // 8 corners of the bounding box
                double[][] corners =
                {
                    new[] { xMin, yMin, zMin },
                    new[] { xMin, yMin, zMax },
                    new[] { xMin, yMax, zMin },
                    new[] { xMin, yMax, zMax },
                    new[] { xMax, yMin, zMin },
                    new[] { xMax, yMin, zMax },
                    new[] { xMax, yMax, zMin },
                    new[] { xMax, yMax, zMax }
                };

                double faceTMin = double.MaxValue;
                double faceTMax = double.MinValue;

                foreach (double[] p in corners)
                {
                    double[] v =
                    {
                        p[0] - axisOrigin[0],
                        p[1] - axisOrigin[1],
                        p[2] - axisOrigin[2]
                    };

                    double t = Dot(v, axisDirUnit);
                    if (t < faceTMin) faceTMin = t;
                    if (t > faceTMax) faceTMax = t;
                }

                if (!any)
                {
                    tMin = faceTMin;
                    tMax = faceTMax;
                    any = true;
                }
                else
                {
                    if (faceTMin < tMin) tMin = faceTMin;
                    if (faceTMax > tMax) tMax = faceTMax;
                }
            }

            if (!any)
                return;

            stackHeightMm = (tMax - tMin) * 1000.0;
        }

        /// <summary>
        /// Pick screw with:
        ///   - largest diameter not exceeding hole diameter (plus small slack)
        ///   - length >= requiredLength if possible, otherwise the longest in that diameter group.
        /// </summary>
        private static StdScrewDefinition FindBestScrew(double holeDiameterMm, double requiredLengthMm)
        {
            var candidatesByDia = StdScrewLibrary.Screws
                .Where(s => s.DiameterMm <= holeDiameterMm + DiameterSlackMm)
                .OrderByDescending(s => s.DiameterMm)
                .ToList();

            if (candidatesByDia.Count == 0)
                return null;

            StdScrewDefinition best =
                candidatesByDia
                    .Where(s => s.LengthMm >= requiredLengthMm)
                    .OrderBy(s => s.LengthMm)
                    .FirstOrDefault();

            if (best != null)
                return best;

            // Otherwise, longest screw of the largest allowed diameter
            return candidatesByDia
                .OrderByDescending(s => s.DiameterMm)
                .ThenByDescending(s => s.LengthMm)
                .First();
        }

        /// <summary>
        /// Add concentric mate between hole cylinder and screw shank.
        /// </summary>
        private static void AddConcentricMate(
            IModelDoc2 swModel,
            IAssemblyDoc asmDoc,
            Component2 screwComp,
            HoleStack stack,
            StdScrewDefinition screw)
        {
            try
            {
                IFace2 holeFace = stack.GetReferenceCylindricalFace();
                IFace2 screwFace = FindCylindricalFaceForScrew(screwComp, screw);

                if (holeFace == null || screwFace == null)
                    return;

                swModel.ClearSelection2(true);

                IEntity entHole = (IEntity)holeFace;
                IEntity entScrew = (IEntity)screwFace;

                bool sel1 = entHole.Select4(false, null);
                bool sel2 = entScrew.Select4(true, null);

                if (!sel1 || !sel2)
                {
                    swModel.ClearSelection2(true);
                    return;
                }

                int mateError;
                Mate2 mate = asmDoc.AddMate5(
                    (int)swMateType_e.swMateCONCENTRIC,
                    (int)swMateAlign_e.swMateAlignALIGNED,
                    false,
                    0, 0, 0, 0, 0, 0, 0, 0,
                    false,
                    false,
                    0,
                    out mateError);

                swModel.ClearSelection2(true);

                if (mateError != (int)swAddMateError_e.swAddMateError_NoError)
                {
                    Debug.WriteLine("AddMate5 returned error: " + mateError);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to add concentric mate: " + ex);
            }
        }

        /// <summary>
        /// Find a cylindrical face in the screw component whose radius matches the screw nominal.
        /// </summary>
        private static IFace2 FindCylindricalFaceForScrew(Component2 comp, StdScrewDefinition screw)
        {
            object bodyInfo;
            object[] bodies = comp.GetBodies3((int)swBodyType_e.swAllBodies, out bodyInfo) as object[];
            if (bodies == null || bodies.Length == 0)
                return null;

            double targetRadius = screw.DiameterMm / 2.0 / 1000.0;

            IFace2 bestFace = null;
            double bestDiff = double.MaxValue;

            foreach (object bodyObj in bodies)
            {
                IBody2 body = bodyObj as IBody2;
                if (body == null)
                    continue;

                object[] faces = body.GetFaces() as object[];
                if (faces == null)
                    continue;

                foreach (object faceObj in faces)
                {
                    IFace2 face = faceObj as IFace2;
                    if (face == null)
                        continue;

                    ISurface surf = face.IGetSurface();
                    if (surf == null || !surf.IsCylinder())
                        continue;

                    double[] cyl = surf.CylinderParams as double[];
                    if (cyl == null || cyl.Length < 7)
                        continue;

                    double radius = Math.Abs(cyl[6]);
                    double diff = Math.Abs(radius - targetRadius);

                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        bestFace = face;
                    }
                }
            }

            return bestFace;
        }

        /// <summary>
        /// Check if two cylinders are coaxial (same axis line in space, within a small tolerance).
        /// </summary>
        private static bool AreCoaxial(HoleFaceInfo a, HoleFaceInfo b)
        {
            double[] v1 = (double[])a.AxisUnit.Clone();
            double[] v2 = (double[])b.AxisUnit.Clone();

            Normalize(v1);
            Normalize(v2);

            double dot = Dot(v1, v2);
            if (dot < 0.0)
            {
                v2[0] = -v2[0];
                v2[1] = -v2[1];
                v2[2] = -v2[2];
                dot = -dot;
            }

            // Axes must be nearly parallel
            if (dot < 0.999) // ~2.5°
                return false;

            // Distance between the two axes
            double[] w0 =
            {
                b.Origin[0] - a.Origin[0],
                b.Origin[1] - a.Origin[1],
                b.Origin[2] - a.Origin[2]
            };

            double proj = Dot(w0, v1);
            double[] wPerp =
            {
                w0[0] - proj * v1[0],
                w0[1] - proj * v1[1],
                w0[2] - proj * v1[2]
            };

            double dist = Norm(wPerp);
            return dist < AxisDistanceTolerance;
        }

        private static double Dot(double[] a, double[] b)
        {
            return a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        }

        private static double Norm(double[] v)
        {
            return Math.Sqrt(Dot(v, v));
        }

        private static void Normalize(double[] v)
        {
            double n = Norm(v);
            if (n < 1e-12)
                return;

            v[0] /= n;
            v[1] /= n;
            v[2] /= n;
        }

        /// <summary>
        /// Ensure that the screw model is loaded in memory so AddComponent5 can use it.
        /// Returns the loaded IModelDoc2 or null on failure.
        /// </summary>
        private static IModelDoc2 EnsureScrewModelLoaded(SldWorks swApp, string filePath,
            out int errors, out int warnings)
        {
            errors = 0;
            warnings = 0;

            if (string.IsNullOrWhiteSpace(filePath))
                return null;

            // Already open? (try full path first)
            IModelDoc2 doc = swApp.GetOpenDocumentByName(filePath) as IModelDoc2;
            if (doc != null)
                return doc;

            // Try by file name only (in case user has it open from another folder)
            string fileNameOnly = Path.GetFileName(filePath);
            if (!string.IsNullOrEmpty(fileNameOnly))
            {
                doc = swApp.GetOpenDocumentByName(fileNameOnly) as IModelDoc2;
                if (doc != null)
                    return doc;
            }

            // Open invisibly as a PART
            bool wasVisible = swApp.GetDocumentVisible((int)swDocumentTypes_e.swDocPART);
            swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);

            doc = swApp.OpenDoc6(
                filePath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors,
                ref warnings) as IModelDoc2;

            // Restore visibility flag
            swApp.DocumentVisible(wasVisible, (int)swDocumentTypes_e.swDocPART);

            return doc;
        }

    }
}
