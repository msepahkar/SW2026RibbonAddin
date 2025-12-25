using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using Clipper2Lib;
using CSMath;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal static partial class DwgLaserNester
    {
        // Geometry scale for Clipper (mm -> integer)
        private const long SCALE = 1000; // 0.001 mm units
        private static readonly int[] RotationsDeg = { 0, 90, 180, 270 };

        // ============================
        // DWG output styling (layers/blocks)
        // ============================

        // Layering makes it easy to toggle visibility (sheets vs info text vs parts).
        private const string LAYER_NEST_SHEETS = "NEST_SHEETS";
        private const string LAYER_NEST_PARTS = "NEST_PARTS";
        private const string LAYER_NEST_BOTTOM_BAR = "NEST_BOTTOM_BAR";
        private const string LAYER_NEST_BOTTOM_TEXT = "NEST_BOTTOM_TEXT";

        // Bottom info bar height (mm) drawn under each sheet.
        private const double NEST_INFO_BAR_HEIGHT_MM = 55.0;

        [ThreadStatic] private static Layer _layerNestSheets;
        [ThreadStatic] private static Layer _layerNestParts;
        [ThreadStatic] private static Layer _layerNestBottomBar;
        [ThreadStatic] private static Layer _layerNestBottomText;

        [ThreadStatic] private static BlockRecord _sheetOutlineBlock;

        // (intentionally left blank)

        // ============================
        // CSV-based material index
        // ============================

        private sealed class AllPartsIndex
        {
            public readonly Dictionary<string, HashSet<string>> MaterialsByThicknessKey =
                new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

            public readonly Dictionary<string, string> MaterialByPartKeyAndThickness =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class AllPartsIndexCacheEntry
        {
            public string CsvPath;
            public DateTime LastWriteTimeUtc;
            public AllPartsIndex Index;
        }

        private static readonly object _indexLock = new object();
        private static readonly Dictionary<string, AllPartsIndexCacheEntry> _allPartsIndexCache =
            new Dictionary<string, AllPartsIndexCacheEntry>(StringComparer.OrdinalIgnoreCase);

        [ThreadStatic] private static AllPartsIndex _activeAllPartsIndex;
        [ThreadStatic] private static string _activeThicknessKey;

        private static string ThicknessKey(double thicknessMm)
        {
            if (thicknessMm <= 0) return "?";
            return thicknessMm.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string FindAllPartsCsv(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder)) return null;

            string[] candidates =
            {
                Path.Combine(folder, "all_parts.csv"),
                Path.Combine(folder, "_all_parts.csv"),
                Path.Combine(folder, "parts.csv"),
                Path.Combine(folder, "_parts.csv"),
            };

            foreach (var p in candidates)
            {
                if (File.Exists(p))
                    return p;
            }

            return null;
        }

        private static AllPartsIndex TryGetAllPartsIndexForFolder(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
                return null;

            string full;
            try { full = Path.GetFullPath(folder); }
            catch { full = folder; }

            // IMPORTANT:
            // 1) Do NOT negative-cache "not found" forever. all_parts.csv may be created later (e.g., after Combine).
            // 2) If the CSV changes (timestamp), reload the index.
            var csvPath = FindAllPartsCsv(full);
            if (csvPath == null)
            {
                lock (_indexLock)
                    _allPartsIndexCache.Remove(full);
                return null;
            }

            DateTime lastWriteUtc;
            try { lastWriteUtc = File.GetLastWriteTimeUtc(csvPath); }
            catch { lastWriteUtc = DateTime.MinValue; }

            lock (_indexLock)
            {
                if (_allPartsIndexCache.TryGetValue(full, out var entry) && entry != null)
                {
                    if (string.Equals(entry.CsvPath, csvPath, StringComparison.OrdinalIgnoreCase) &&
                        entry.LastWriteTimeUtc == lastWriteUtc)
                    {
                        return entry.Index;
                    }
                }
            }

            AllPartsIndex idx;
            try
            {
                idx = LoadAllPartsIndex(csvPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("LoadAllPartsIndex failed: " + ex);
                idx = null;
            }

            lock (_indexLock)
            {
                _allPartsIndexCache[full] = new AllPartsIndexCacheEntry
                {
                    CsvPath = csvPath,
                    LastWriteTimeUtc = lastWriteUtc,
                    Index = idx
                };
            }

            return idx;
        }

        internal static void ClearAllPartsIndexCache()
        {
            lock (_indexLock)
                _allPartsIndexCache.Clear();
        }

        internal static void ClearAllPartsIndexCacheForFolder(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
                return;

            string full;
            try { full = Path.GetFullPath(folder); }
            catch { full = folder; }

            lock (_indexLock)
                _allPartsIndexCache.Remove(full);
        }

        private static AllPartsIndex LoadAllPartsIndex(string csvPath)
        {
            var idx = new AllPartsIndex();

            using (var sr = new StreamReader(csvPath, Encoding.UTF8, true))
            {
                string header = sr.ReadLine();
                if (header == null)
                    return idx;

                header = header.TrimStart('\uFEFF');
                var headerFields = SplitCsvLine(header);

                int colFile = FindCol(headerFields, "FileName", "Filename", "File");
                int colThk = FindCol(headerFields, "PlateThickness_mm", "Thickness", "PlateThickness");
                int colMat = FindCol(headerFields,
                    "Material",
                    "SWMaterial",
                    "SW-Material",
                    "SolidWorksMaterial",
                    "MaterialExact",
                    "MaterialName");

                if (colFile < 0 || colThk < 0 || colMat < 0)
                    return idx;

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var fields = SplitCsvLine(line);
                    if (fields.Count <= Math.Max(colMat, Math.Max(colFile, colThk)))
                        continue;

                    string fileName = (fields[colFile] ?? "").Trim();
                    string thkStr = (fields[colThk] ?? "").Trim();
                    string material = (fields[colMat] ?? "").Trim();

                    if (fileName.Length == 0)
                        continue;

                    if (!double.TryParse(thkStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double thk) || thk <= 0)
                        continue;

                    material = NormalizeMaterialLabel(material);
                    string thkKey = ThicknessKey(thk);

                    if (!idx.MaterialsByThicknessKey.TryGetValue(thkKey, out var set))
                    {
                        set = new HashSet<string>(StringComparer.Ordinal);
                        idx.MaterialsByThicknessKey[thkKey] = set;
                    }
                    set.Add(material);

                    string baseName = Path.GetFileNameWithoutExtension(fileName) ?? fileName;
                    string partLoose = MakeLooseKey(baseName);

                    string key = partLoose + "|" + thkKey;
                    if (!idx.MaterialByPartKeyAndThickness.ContainsKey(key))
                        idx.MaterialByPartKeyAndThickness[key] = material;
                }
            }

            return idx;
        }

        private static int FindCol(List<string> headerFields, params string[] names)
        {
            if (headerFields == null || headerFields.Count == 0 || names == null || names.Length == 0)
                return -1;

            string Norm(string s)
            {
                s = (s ?? "").Trim().Trim('"').Trim();
                if (s.Length == 0) return "";
                s = s.Replace(" ", "").Replace("-", "").Replace("_", "");
                return s.ToUpperInvariant();
            }

            // Normalize header fields once.
            for (int i = 0; i < headerFields.Count; i++)
            {
                var h = Norm(headerFields[i]);
                if (h.Length == 0) continue;

                foreach (var nRaw in names)
                {
                    var n = Norm(nRaw);
                    if (n.Length == 0) continue;

                    if (h == n)
                        return i;
                }
            }

            return -1;
        }

        private static List<string> SplitCsvLine(string line)
        {
            var res = new List<string>();
            if (line == null) return res;

            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    res.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            res.Add(sb.ToString());
            return res;
        }

        private static string MakeLooseKey(string s)
        {
            s = (s ?? "").Trim().ToUpperInvariant();
            if (s.Length == 0) return "";

            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);
            }
            return sb.ToString();
        }

        private static bool TryExtractPartTokenFromBlockName(string blockName, out string token)
        {
            token = null;
            if (string.IsNullOrWhiteSpace(blockName))
                return false;

            if (!blockName.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                return false;

            int qIndex = blockName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            if (qIndex < 0)
                return false;

            int start = 2;
            int len = qIndex - start;
            if (len <= 0)
                return false;

            token = blockName.Substring(start, len);
            token = (token ?? "").Trim();
            if (token.Length == 0)
                return false;

            // If the block name includes an embedded material token (e.g. __MATB64_<token>__),
            // strip it out so CSV-based lookups (keyed by the original part base name) still work.
            token = StripEmbeddedMaterialToken(token);

            return token.Length > 0;
        }

        private static string StripEmbeddedMaterialToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
                return token;

            // New reversible token used by the combiner.
            token = RemoveTokenSegment(token, MaterialNameCodec.BlockTokenPrefix, MaterialNameCodec.BlockTokenSuffix);

            // Legacy (pre-b64) token formats the nester already supported.
            token = RemoveTokenSegment(token, "__MAT_", "__");

            // Cosmetic cleanup (helps MakeLooseKey match better).
            while (token.Contains("__"))
                token = token.Replace("__", "_");

            return token.Trim('_', ' ');
        }

        private static string RemoveTokenSegment(string s, string prefix, string suffix)
        {
            if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(prefix) || string.IsNullOrEmpty(suffix))
                return s;

            int p = s.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (p < 0)
                return s;

            int start = p;
            int afterPrefix = p + prefix.Length;
            int end = s.IndexOf(suffix, afterPrefix, StringComparison.OrdinalIgnoreCase);

            // If we can't find the closing suffix, strip to end (defensive against truncated names)
            if (end >= 0)
                end += suffix.Length;
            else
                end = s.Length;

            try
            {
                return s.Remove(start, end - start);
            }
            catch
            {
                return s;
            }
        }

        // ============================
        // Types
        // ============================

        internal sealed class NestRunResult
        {
            public string ThicknessFile;
            public string MaterialExact;
            public string OutputDwg;
            public int SheetsUsed;
            public int TotalParts;
            public NestingMode Mode;
        }

        private sealed class PartDefinition
        {
            public BlockRecord Block;
            public string BlockName;
            public int Quantity;
            public int PartCountWeight = 1;

            public string MaterialExact;

            public double MinX, MinY, MaxX, MaxY;
            public double Width, Height;

            public Path64 OuterContour0;
            public long OuterArea2Abs;
        }

        /// <summary>
        /// Returns the best available estimate of the TRUE part material area (in "2x area" units).
        /// We keep this separate from the placement polygon because:
        ///  - Fast(Rectangles) packs using bounding boxes
        ///  - Mirror-pair blocks may pack using a rectangle envelope (especially in Gap mode)
        /// but users still want sheet fill based on real part area.
        /// </summary>
        private static long GetRealArea2Abs(PartDefinition part)
        {
            if (part == null)
                return 0;

            if (part.OuterArea2Abs > 0)
                return part.OuterArea2Abs;

            // Fallback: bounding rectangle area.
            long w = ToInt(Math.Max(0.0, part.Width));
            long h = ToInt(Math.Max(0.0, part.Height));
            if (w <= 0 || h <= 0)
                return 0;

            return 2L * w * h;
        }

        private sealed class FreeRect
        {
            public double X, Y, W, H;
        }

        private sealed class SheetRectState
        {
            public int Index;
            public double OriginXmm;
            public double OriginYmm;
            public List<FreeRect> Free = new List<FreeRect>();

            // For reporting (Fill %)
            public int PlacedCount;
            public long UsedArea2Abs;
        }

        private sealed class SheetContourState
        {
            public int Index;
            public double OriginXmm;
            public double OriginYmm;

            public List<PlacedContour> Placed = new List<PlacedContour>();

            public int PlacedCount;
            public long UsedArea2Abs;
        }

        private sealed class PlacedContour
        {
            // Full precision (used for collision checks)
            public Path64 OffsetPoly;

            // Cached decimated version for NFP/Minkowski (speed)
            public Path64 OffsetPolyNfp;

            public LongRect BBox;
        }

        private struct LongRect
        {
            public long MinX, MinY, MaxX, MaxY;
        }

        private struct CandidateIns
        {
            public long InsX, InsY;
        }

        private struct RotatedPoly
        {
            public int RotDeg;
            public double RotRad;

            public Path64 PolyRot;     // rotated (no gap)
            public Path64 PolyOffset;  // rotated + gap/2 offset (full precision)

            // Cached NFP versions (decimated + negated)
            public Path64 PolyOffsetNfp;
            public Path64 PolyOffsetNfpNeg;

            public LongRect OffsetBounds;
            public Point64[] Anchors;

            public long RotArea2Abs;
        }

        private struct ContourPlacement
        {
            public long InsertX;
            public long InsertY;
            public double RotRad;

            public Path64 OffsetPolyTranslated;
            public LongRect OffsetBBoxTranslated;

            public long RotArea2Abs;
        }

        // ============================
        // Scan jobs
        // ============================
        public static List<LaserNestJob> ScanJobsForFolder(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return new List<LaserNestJob>();

            var thicknessFiles = GetThicknessFiles(folder);
            var jobs = new List<LaserNestJob>();

            var idx = TryGetAllPartsIndexForFolder(folder);

            foreach (var file in thicknessFiles)
            {
                double thickness = TryGetPlateThicknessFromFileName(file) ?? 0.0;
                string thkKey = ThicknessKey(thickness);

                // Preferred: use all_parts.csv (fast)
                HashSet<string> mats = null;
                if (idx != null && thickness > 0 && idx.MaterialsByThicknessKey.TryGetValue(thkKey, out var set) && set != null && set.Count > 0)
                {
                    mats = new HashSet<string>(set, StringComparer.Ordinal);

                    // If we have known materials, drop UNKNOWN to avoid clutter.
                    if (mats.Count > 1)
                        mats.RemoveWhere(m => string.Equals(m, "UNKNOWN", StringComparison.OrdinalIgnoreCase));
                }

                // Fallback: scan the DWG itself for embedded material tokens (robust even if CSV missing)
                if (mats == null || mats.Count == 0)
                    mats = TryScanMaterialsFromThicknessDwg(file, idx, thkKey);

                if (mats != null && mats.Count > 0)
                {
                    foreach (var mat in mats.OrderBy(s => s, StringComparer.Ordinal))
                    {
                        jobs.Add(new LaserNestJob
                        {
                            Enabled = true,
                            ThicknessFilePath = file,
                            ThicknessMm = thickness,
                            MaterialExact = mat
                        });
                    }
                }
                else
                {
                    // Last resort: unknown
                    jobs.Add(new LaserNestJob
                    {
                        Enabled = true,
                        ThicknessFilePath = file,
                        ThicknessMm = thickness,
                        MaterialExact = "UNKNOWN"
                    });
                }
            }

            return jobs
                .OrderBy(j => j.MaterialExact ?? "", StringComparer.Ordinal)
                .ThenBy(j => j.ThicknessMm <= 0 ? double.MaxValue : j.ThicknessMm)
                .ThenBy(j => j.ThicknessFileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static HashSet<string> TryScanMaterialsFromThicknessDwg(string dwgPath, AllPartsIndex idx, string thicknessKey)
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
                return null;

            try
            {
                CadDocument doc;
                using (var reader = new DwgReader(dwgPath))
                {
                    doc = reader.Read();
                }

                if (doc == null)
                    return null;

                var materials = new HashSet<string>(StringComparer.Ordinal);

                foreach (var block in doc.BlockRecords)
                {
                    if (block == null)
                        continue;

                    string name = block.Name ?? "";
                    if (!name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Only consider our "plate" blocks (they carry Qty + material token)
                    if (name.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    string material = "UNKNOWN";

                    bool gotMatFromName = TryExtractMaterialFromBlockName(name, out var matFromName);
                    if (gotMatFromName)
                        material = NormalizeMaterialLabel(matFromName);

                    bool isUnknown = string.Equals(material, "UNKNOWN", StringComparison.OrdinalIgnoreCase);

                    // If the block-name material is missing/unknown, try CSV mapping by (partKey + thickness)
                    if ((!gotMatFromName || isUnknown) && idx != null && !string.IsNullOrWhiteSpace(thicknessKey))
                    {
                        if (TryExtractPartTokenFromBlockName(name, out var token))
                        {
                            string partLoose = MakeLooseKey(token);
                            string key = partLoose + "|" + thicknessKey;

                            if (idx.MaterialByPartKeyAndThickness.TryGetValue(key, out var matFromCsv))
                                material = NormalizeMaterialLabel(matFromCsv);
                        }
                    }

                    material = NormalizeMaterialLabel(material);
                    materials.Add(material);
                }

                if (materials.Count > 1)
                    materials.RemoveWhere(m => string.Equals(m, "UNKNOWN", StringComparison.OrdinalIgnoreCase));

                return materials.Count > 0 ? materials : null;
            }
            catch
            {
                return null;
            }
        }

        private static List<string> GetThicknessFiles(string folder)
        {
            return Directory.GetFiles(folder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
                .Where(f =>
                {
                    string n = Path.GetFileNameWithoutExtension(f) ?? "";
                    bool isNested = n.IndexOf("_nested", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool isNestLog = n.IndexOf("_nest_", StringComparison.OrdinalIgnoreCase) >= 0;
                    return !isNested && !isNestLog;
                })
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ============================
        // Batch nesting (CANCEL-aware)
        // ============================
        public static void NestJobs(string folder, List<LaserNestJob> jobs, LaserCutRunSettings settings, bool showUi = true)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));
            if (jobs == null)
                throw new ArgumentNullException(nameof(jobs));

            var selected = jobs.Where(j => j != null && j.Enabled).ToList();
            if (selected.Count == 0)
                return;

            NestBatchResult run;

            using (var progress = new LaserCutProgressForm())
            {
                progress.Show();
                run = NestJobsWithProgress(folder, selected, settings, progress, showUi: false);
                try { progress.Close(); } catch { }
            }

            if (showUi)
            {
                MessageBox.Show(
                    run.Cancelled
                        ? "Nesting was cancelled.\r\n\r\nSummary:\r\n" + run.SummaryPath
                        : run.FailedTasks > 0
                            ? $"Batch nesting finished with {run.FailedTasks} failed task(s).\r\n\r\nSummary:\r\n" + run.SummaryPath
                            : "Batch nesting finished.\r\n\r\nSummary:\r\n" + run.SummaryPath,
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    run.Cancelled || run.FailedTasks > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        internal sealed class NestBatchResult
        {
            public string SummaryPath;
            public int TotalTasks;
            public int DoneTasks;
            public int FailedTasks;
            public bool Cancelled;
        }

        internal static NestBatchResult NestJobsWithProgress(
            string folder,
            List<LaserNestJob> jobs,
            LaserCutRunSettings settings,
            ILaserCutProgress progress,
            bool showUi = false)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));
            if (jobs == null)
                throw new ArgumentNullException(nameof(jobs));

            progress ??= new NullLaserCutProgress();

            var selected = jobs.Where(j => j != null && j.Enabled).ToList();
            var result = new NestBatchResult
            {
                SummaryPath = Path.Combine(folder ?? "", "batch_nest_summary.txt"),
                TotalTasks = selected.Count,
                DoneTasks = 0,
                FailedTasks = 0,
                Cancelled = false
            };

            progress.BeginBatch(selected.Count);

            // Build summary header
            var summary = new StringBuilder();
            summary.AppendLine("Batch nesting summary");
            summary.AppendLine("Folder: " + folder);
            summary.AppendLine("Tasks: " + selected.Count);
            summary.AppendLine("Mode: " + settings.Mode);
            summary.AppendLine("Note: SeparateByMaterialExact / OneDWGPerMaterial / FilterPreview are FORCED true.");
            summary.AppendLine(new string('-', 70));

            for (int i = 0; i < selected.Count; i++)
            {
                var job = selected[i];
                int taskIndex = i + 1;

                try
                {
                    progress.ThrowIfCancelled();

                    var run = RunSingleJob(job, settings, progress, taskIndex, selected.Count);
                    result.DoneTasks++;

                    summary.AppendLine(Path.GetFileName(run.ThicknessFile));
                    summary.AppendLine($"  Material: {run.MaterialExact}");
                    summary.AppendLine($"  Mode: {run.Mode}");
                    summary.AppendLine($"  SheetsUsed: {run.SheetsUsed}, Parts: {run.TotalParts}");
                    summary.AppendLine($"  Output: {Path.GetFileName(run.OutputDwg)}");
                    summary.AppendLine(new string('-', 70));

                    progress.EndTask(result.DoneTasks + result.FailedTasks, selected.Count, job, success: true,
                        message: $"SheetsUsed: {run.SheetsUsed}, Parts: {run.TotalParts}");
                }
                catch (OperationCanceledException)
                {
                    result.Cancelled = true;
                    summary.AppendLine("CANCELLED BY USER");
                    summary.AppendLine(new string('-', 70));

                    // Mark current task as cancelled/failed for UI purposes.
                    progress.EndTask(result.DoneTasks + result.FailedTasks, selected.Count, job, success: false, message: "Cancelled");
                    break;
                }
                catch (Exception ex)
                {
                    result.FailedTasks++;

                    summary.AppendLine(Path.GetFileName(job.ThicknessFilePath ?? "(missing file)"));
                    summary.AppendLine($"  Material: {NormalizeMaterialLabel(job.MaterialExact)}");
                    summary.AppendLine("  ERROR:");
                    summary.AppendLine("  " + ex.Message);
                    summary.AppendLine(new string('-', 70));

                    progress.EndTask(result.DoneTasks + result.FailedTasks, selected.Count, job, success: false, message: ex.Message);

                    // Continue with remaining jobs.
                }
            }

            try
            {
                File.WriteAllText(result.SummaryPath, summary.ToString(), Encoding.UTF8);
            }
            catch { }

            return result;
        }

        private sealed class NullLaserCutProgress : ILaserCutProgress
        {
            public void BeginBatch(int totalTasks) { }
            public void BeginTask(int taskIndex, int totalTasks, LaserNestJob job, int totalParts, NestingMode mode, double sheetWmm, double sheetHmm) { }
            public void ReportPlaced(int placed, int total, int sheetsUsed) { }
            public void EndTask(int doneTasks, int totalTasks, LaserNestJob job, bool success, string message) { }
            public void SetStatus(string message) { }
            public void ThrowIfCancelled() { }
        }

        // ============================
        // Single job
        // ============================
        private static NestRunResult RunSingleJob(LaserNestJob job, LaserCutRunSettings settings, ILaserCutProgress progress, int taskIndex, int totalTasks)
        {
            progress.ThrowIfCancelled();

            if (job == null)
                throw new ArgumentNullException(nameof(job));

            if (!File.Exists(job.ThicknessFilePath))
                throw new FileNotFoundException("DWG file not found.", job.ThicknessFilePath);

            if (job.Sheet.WidthMm <= 0 || job.Sheet.HeightMm <= 0)
                throw new InvalidOperationException("Invalid sheet size for job: " + job.MaterialExact);

            string thicknessFileName = Path.GetFileName(job.ThicknessFilePath) ?? job.ThicknessFilePath;

            double thicknessMm = job.ThicknessMm > 0 ? job.ThicknessMm : (TryGetPlateThicknessFromFileName(job.ThicknessFilePath) ?? 0.0);
            string thkKey = ThicknessKey(thicknessMm);

            string material = NormalizeMaterialLabel(job.MaterialExact);

            double gapMm = 3.0;
            if (thicknessMm > gapMm) gapMm = thicknessMm;

            double marginMm = 10.0;
            if (thicknessMm > marginMm) marginMm = thicknessMm;

            // Load fresh
            CadDocument doc;
            using (var reader = new DwgReader(job.ThicknessFilePath))
                doc = reader.Read();

            // activate CSV mapping for this thickness while scanning blocks
            var idx = TryGetAllPartsIndexForFolder(Path.GetDirectoryName(job.ThicknessFilePath) ?? "");
            var prevIdx = _activeAllPartsIndex;
            var prevKey = _activeThicknessKey;
            _activeAllPartsIndex = idx;
            _activeThicknessKey = thkKey;

            List<PartDefinition> defsAll;
            try
            {
                defsAll = LoadPartDefinitions(doc, settings).ToList();
            }
            finally
            {
                _activeAllPartsIndex = prevIdx;
                _activeThicknessKey = prevKey;
            }

            var defs = defsAll
                .Where(d => string.Equals(NormalizeMaterialLabel(d.MaterialExact), material, StringComparison.Ordinal))
                .ToList();

            int totalParts = defs.Sum(d => d.Quantity * Math.Max(1, d.PartCountWeight));

            // if mapping failed, fallback to all (still produce output)
            if (totalParts <= 0)
            {
                defs = defsAll;
                totalParts = defs.Sum(d => d.Quantity * Math.Max(1, d.PartCountWeight));
            }

            BlockRecord modelSpace = doc.BlockRecords["*Model_Space"];

            // keep only that material's preview (forced true)
            if (defs.Count > 0)
            {
                var keepSet = new HashSet<string>(defs.Select(d => d.BlockName), StringComparer.OrdinalIgnoreCase);
                FilterSourcePreviewToTheseBlocks(doc, keepSet);
            }

            GetModelSpaceExtents(doc, out double srcMinX, out _, out _, out double srcMaxY);

            double baseSheetOriginX = srcMinX;
            double baseSheetOriginY = srcMaxY + 200.0;
            // Prepare DWG output layers/blocks for this job.
            EnsureNestOutputLayersAndBlocks(doc, job.Sheet.WidthMm, job.Sheet.HeightMm);

            progress.BeginTask(
                taskIndex,
                totalTasks,
                job,
                totalParts,
                settings.Mode,
                job.Sheet.WidthMm,
                job.Sheet.HeightMm);

            int sheetsUsed;
            NestingMode finalMode = settings.Mode;

            try
            {
                if (settings.Mode == NestingMode.FastRectangles)
                {
                    sheetsUsed = NestFastRectangles(
                        defs,
                        modelSpace,
                        job.Sheet.WidthMm,
                        job.Sheet.HeightMm,
                        marginMm,
                        gapMm,
                        baseSheetOriginX,
                        baseSheetOriginY,
                        material,
                        progress,
                        totalParts);
                }
                else if (settings.Mode == NestingMode.ContourLevel1)
                {
                    sheetsUsed = NestContourLevel1(
                        defs,
                        modelSpace,
                        job.Sheet.WidthMm,
                        job.Sheet.HeightMm,
                        marginMm,
                        gapMm,
                        baseSheetOriginX,
                        baseSheetOriginY,
                        material,
                        progress,
                        totalParts,
                        chordMm: Math.Max(0.10, settings.ContourChordMm),
                        snapMm: Math.Max(0.01, settings.ContourSnapMm),
                        maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry));
                }
                else
                {
                    // NOTE: Level 2 is now HYBRID inside the algorithm implementation
                    sheetsUsed = NestContourLevel2_Nfp(
                        defs,
                        modelSpace,
                        job.Sheet.WidthMm,
                        job.Sheet.HeightMm,
                        marginMm,
                        gapMm,
                        baseSheetOriginX,
                        baseSheetOriginY,
                        material,
                        progress,
                        totalParts,
                        chordMm: Math.Max(0.10, settings.ContourChordMm),
                        snapMm: Math.Max(0.01, settings.ContourSnapMm),
                        maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry),
                        maxPartners: Math.Max(10, settings.MaxNfpPartnersPerTry));
                }
            }
            catch (Level2TimeoutException)
            {
                // If Level2 still times out, fallback to Level1 for this job
                finalMode = NestingMode.ContourLevel1;
                progress.SetStatus("Level 2 timeout → falling back to Level 1 for this job...");

                // reload fresh doc
                using (var reader = new DwgReader(job.ThicknessFilePath))
                    doc = reader.Read();

                modelSpace = doc.BlockRecords["*Model_Space"];

                if (defs.Count > 0)
                {
                    var keepSet = new HashSet<string>(defs.Select(d => d.BlockName), StringComparer.OrdinalIgnoreCase);
                    FilterSourcePreviewToTheseBlocks(doc, keepSet);
                }

                GetModelSpaceExtents(doc, out srcMinX, out _, out _, out srcMaxY);
                baseSheetOriginX = srcMinX;
                baseSheetOriginY = srcMaxY + 200.0;

                // Re-create DWG output layers/blocks for the freshly reloaded document.
                EnsureNestOutputLayersAndBlocks(doc, job.Sheet.WidthMm, job.Sheet.HeightMm);

                progress.BeginTask(
                    taskIndex,
                    totalTasks,
                    job,
                    totalParts,
                    finalMode,
                    job.Sheet.WidthMm,
                    job.Sheet.HeightMm);

                sheetsUsed = NestContourLevel1(
                    defs,
                    modelSpace,
                    job.Sheet.WidthMm,
                    job.Sheet.HeightMm,
                    marginMm,
                    gapMm,
                    baseSheetOriginX,
                    baseSheetOriginY,
                    material,
                    progress,
                    totalParts,
                    chordMm: Math.Max(0.10, settings.ContourChordMm),
                    snapMm: Math.Max(0.01, settings.ContourSnapMm),
                    maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry));
            }

            progress.ThrowIfCancelled();

            string dir = Path.GetDirectoryName(job.ThicknessFilePath) ?? "";
            string nameNoExt = Path.GetFileNameWithoutExtension(job.ThicknessFilePath) ?? "thickness";

            string safeMat = MakeSafeFileToken(material);
            string outPath = Path.Combine(dir, $"{nameNoExt}_nested_{safeMat}.dwg");

            using (var writer = new DwgWriter(outPath, doc))
                writer.Write();

            string logPath = Path.Combine(dir, $"{nameNoExt}_nest_log.txt");
            AppendNestLog(
                logPath,
                job.ThicknessFilePath,
                material,
                job.Sheet.WidthMm,
                job.Sheet.HeightMm,
                thicknessMm,
                gapMm,
                marginMm,
                sheetsUsed,
                totalParts,
                outPath,
                finalMode);

            return new NestRunResult
            {
                ThicknessFile = job.ThicknessFilePath,
                MaterialExact = material,
                OutputDwg = outPath,
                SheetsUsed = sheetsUsed,
                TotalParts = totalParts,
                Mode = finalMode
            };
        }

        // ---------------------------
        // Part scanning (blocks)
        // ---------------------------
        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc, LaserCutRunSettings settings)
        {
            if (doc == null)
                yield break;

            foreach (var block in doc.BlockRecords)
            {
                if (block == null) continue;

                string name = block.Name ?? "";
                if (!name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                int qIndex = name.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
                if (qIndex < 0)
                    continue;

                int qty = 1;
                int start = qIndex + 2;
                int end = start;
                while (end < name.Length && char.IsDigit(name[end]))
                    end++;

                if (end > start)
                {
                    string qtyToken = name.Substring(start, end - start);
                    if (!int.TryParse(qtyToken, NumberStyles.Integer, CultureInfo.InvariantCulture, out qty))
                        qty = 1;
                }

                string material = "UNKNOWN";

                bool gotMatFromName = TryExtractMaterialFromBlockName(name, out var matFromName);
                if (gotMatFromName)
                    material = NormalizeMaterialLabel(matFromName);

                // If the block-name material is missing/unknown (e.g. legacy blocks or truncated tokens),
                // fall back to all_parts.csv mapping (base-name + thickness).
                bool isUnknown = string.Equals(material, "UNKNOWN", StringComparison.OrdinalIgnoreCase);

                if ((!gotMatFromName || isUnknown) && _activeAllPartsIndex != null && !string.IsNullOrWhiteSpace(_activeThicknessKey))
                {
                    if (TryExtractPartTokenFromBlockName(name, out var token))
                    {
                        string partLoose = MakeLooseKey(token);
                        string key = partLoose + "|" + _activeThicknessKey;

                        if (_activeAllPartsIndex.MaterialByPartKeyAndThickness.TryGetValue(key, out var matFromCsv))
                            material = NormalizeMaterialLabel(matFromCsv);
                    }
                }

                if (!TryGetBlockBbox(block, out double minX, out double minY, out double maxX, out double maxY))
                    continue;

                double w = maxX - minX;
                double h = maxY - minY;
                if (w <= 0 || h <= 0)
                    continue;

                Path64 contour = null;
                long area2Abs = 0;

                try
                {
                    contour = ExtractOuterContourScaled(block, chordMm: settings.ContourChordMm, snapMm: settings.ContourSnapMm);
                    contour = CleanPath(contour);
                    area2Abs = Area2Abs(contour);
                }
                catch
                {
                    contour = null;
                    area2Abs = 0;
                }

                if (contour == null || contour.Count < 3)
                {
                    contour = MakeRectPolyScaled(minX, minY, maxX, maxY);
                    area2Abs = Area2Abs(contour);
                }

                yield return new PartDefinition
                {
                    Block = block,
                    BlockName = name,
                    Quantity = Math.Max(1, qty),
                    MaterialExact = material,

                    MinX = minX,
                    MinY = minY,
                    MaxX = maxX,
                    MaxY = maxY,
                    Width = w,
                    Height = h,

                    OuterContour0 = contour,
                    OuterArea2Abs = area2Abs
                };
            }
        }


        private static bool TryGetBlockBbox(BlockRecord block, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            bool any = false;

            foreach (var ent in block.Entities)
            {
                try
                {
                    var bb = ent.GetBoundingBox();
                    var a = bb.Min;
                    var b = bb.Max;

                    if (a.X < minX) minX = a.X;
                    if (a.Y < minY) minY = a.Y;
                    if (b.X > maxX) maxX = b.X;
                    if (b.Y > maxY) maxY = b.Y;

                    any = true;
                }
                catch { }
            }

            if (!any || minX == double.MaxValue || maxX == double.MinValue)
                return false;

            return true;
        }

        private static string NormalizeMaterialLabel(string s)
        {
            s = (s ?? "").Trim();
            return string.IsNullOrWhiteSpace(s) ? "UNKNOWN" : s;
        }

        private static bool TryExtractMaterialFromBlockName(string blockName, out string material)
        {
            material = null;

            if (string.IsNullOrWhiteSpace(blockName))
                return false;

            // Preferred (new) format: the combiner encodes the *exact* SolidWorks material
            // into the block name using a reversible Base64Url token.
            // Example: P_PART__MATB64_<token>___Q10
            try
            {
                if (MaterialNameCodec.TryExtractFromBlockName(blockName, out var exact))
                {
                    material = exact;
                    return true;
                }
            }
            catch
            {
                // Defensive: fall back to legacy heuristics below.
            }

            string[] markers = new[]
            {
                "__MAT(", "__MAT[", "__MAT=", "__MAT:", "|MAT=", "|MAT:", "_MAT(", "_MAT[", "_MAT=", "_MAT:"
            };

            foreach (var m in markers)
            {
                int idx = blockName.IndexOf(m, StringComparison.OrdinalIgnoreCase);
                if (idx < 0)
                    continue;

                int start = idx + m.Length;

                string token;

                if (m.EndsWith("(", StringComparison.Ordinal))
                {
                    int end = blockName.IndexOf(')', start);
                    if (end < 0) end = blockName.Length;
                    token = blockName.Substring(start, end - start);
                }
                else if (m.EndsWith("[", StringComparison.Ordinal))
                {
                    int end = blockName.IndexOf(']', start);
                    if (end < 0) end = blockName.Length;
                    token = blockName.Substring(start, end - start);
                }
                else
                {
                    int end = blockName.Length;

                    int end1 = blockName.IndexOf("__", start, StringComparison.Ordinal);
                    if (end1 >= 0) end = Math.Min(end, end1);

                    int end2 = blockName.IndexOf("|", start, StringComparison.Ordinal);
                    if (end2 >= 0) end = Math.Min(end, end2);

                    token = blockName.Substring(start, Math.Max(0, end - start));
                }

                token = (token ?? "").Trim();
                if (token.Length == 0)
                    continue;

                try { token = Uri.UnescapeDataString(token); } catch { }

                token = token.Trim();
                if (token.Length == 0)
                    continue;

                material = token;
                return true;
            }

            return false;
        }

        private static double? TryGetPlateThicknessFromFileName(string sourceDwgPath)
        {
            if (string.IsNullOrWhiteSpace(sourceDwgPath))
                return null;

            string fileName = Path.GetFileNameWithoutExtension(sourceDwgPath);
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            const string prefix = "thickness_";
            int idx = fileName.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return null;

            string token = fileName.Substring(idx + prefix.Length);
            if (string.IsNullOrWhiteSpace(token))
                return null;

            token = token.Replace('_', '.');

            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                value > 0.0 && value < 1000.0)
            {
                return value;
            }

            return null;
        }

        // ============================
        // Preview filtering (keeps only selected blocks + nearby text)
        // ============================
        private static void FilterSourcePreviewToTheseBlocks(CadDocument doc, HashSet<string> keepBlockNames)
        {
            if (doc == null || keepBlockNames == null || keepBlockNames.Count == 0)
                return;

            BlockRecord modelSpace;
            try { modelSpace = doc.BlockRecords["*Model_Space"]; }
            catch { return; }

            var inserts = modelSpace.Entities.OfType<Insert>().ToList();

            var keepRanges = new List<(double minX, double maxX)>();

            var defMap = new Dictionary<string, (double minX, double minY, double maxX, double maxY)>(StringComparer.OrdinalIgnoreCase);
            foreach (var br in doc.BlockRecords)
            {
                if (br == null) continue;
                string n = br.Name ?? "";
                if (!n.StartsWith("P_", StringComparison.OrdinalIgnoreCase)) continue;

                if (TryGetBlockBbox(br, out double mnX, out double mnY, out double mxX, out double mxY))
                    defMap[n] = (mnX, mnY, mxX, mxY);
            }

            foreach (var ins in inserts)
            {
                var blk = ins.Block;
                if (blk == null) continue;

                string bn = blk.Name ?? "";
                if (!bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!keepBlockNames.Contains(bn))
                    continue;

                if (!defMap.TryGetValue(bn, out var bb))
                    continue;

                double ix = ins.InsertPoint.X;
                double minX = ix + bb.minX;
                double maxX = ix + bb.maxX;
                if (minX > maxX) { var t = minX; minX = maxX; maxX = t; }
                keepRanges.Add((minX, maxX));
            }

            const double pad = 150.0;

            bool IsNear(double x)
            {
                foreach (var r in keepRanges)
                {
                    if (x >= r.minX - pad && x <= r.maxX + pad)
                        return true;
                }
                return false;
            }

            var remove = new List<Entity>();

            foreach (var e in modelSpace.Entities)
            {
                if (e is Insert ins)
                {
                    var blk = ins.Block;
                    if (blk == null) continue;

                    string bn = blk.Name ?? "";
                    if (bn.StartsWith("P_", StringComparison.OrdinalIgnoreCase) && !keepBlockNames.Contains(bn))
                        remove.Add(e);
                }
                else if (e is MText mt)
                {
                    if (!IsNear(mt.InsertPoint.X))
                        remove.Add(e);
                }
            }

            foreach (var e in remove.Distinct())
            {
                try { modelSpace.Entities.Remove(e); } catch { }
            }
        }

        private static void GetModelSpaceExtents(CadDocument doc, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            BlockRecord modelSpace;
            try { modelSpace = doc.BlockRecords["*Model_Space"]; }
            catch
            {
                minX = minY = maxX = maxY = 0.0;
                return;
            }

            bool any = false;

            foreach (var ent in modelSpace.Entities)
            {
                try
                {
                    var bb = ent.GetBoundingBox();
                    var a = bb.Min;
                    var b = bb.Max;

                    if (a.X < minX) minX = a.X;
                    if (a.Y < minY) minY = a.Y;
                    if (b.X > maxX) maxX = b.X;
                    if (b.Y > maxY) maxY = b.Y;

                    any = true;
                }
                catch { }
            }

            if (!any || minX == double.MaxValue || maxX == double.MinValue)
                minX = minY = maxX = maxY = 0.0;
        }

        private static void EnsureNestOutputLayersAndBlocks(CadDocument doc, double sheetWmm, double sheetHmm)
        {
            if (doc == null)
                return;

            // Layers
            _layerNestSheets = GetOrCreateLayer(doc, LAYER_NEST_SHEETS);
            _layerNestParts = GetOrCreateLayer(doc, LAYER_NEST_PARTS);
            _layerNestBottomBar = GetOrCreateLayer(doc, LAYER_NEST_BOTTOM_BAR);
            _layerNestBottomText = GetOrCreateLayer(doc, LAYER_NEST_BOTTOM_TEXT);

            // Blocks
            _sheetOutlineBlock = GetOrCreateSheetOutlineBlock(doc, sheetWmm, sheetHmm);
        }

        private static Layer GetOrCreateLayer(CadDocument doc, string layerName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(layerName))
                return null;

            try
            {
                return doc.Layers[layerName];
            }
            catch { }

            try
            {
                var layer = new Layer(layerName);
                doc.Layers.Add(layer);
                return layer;
            }
            catch
            {
                try { return doc.Layers[layerName]; } catch { return null; }
            }
        }

        private static string MakeSafeBlockName(string s)
        {
            s = MakeSafeFileToken(s);
            s = (s ?? "BLOCK").Replace('.', 'p').Replace('-', '_');
            if (s.Length > 120)
                s = s.Substring(0, 120);
            return string.IsNullOrWhiteSpace(s) ? "BLOCK" : s;
        }

        private static BlockRecord GetOrCreateSheetOutlineBlock(CadDocument doc, double sheetWmm, double sheetHmm)
        {
            if (doc == null)
                return null;

            // One block per (job sheet size) keeps the model clean.
            string name = MakeSafeBlockName($"NEST_SHEET_{sheetWmm:0.###}x{sheetHmm:0.###}mm");

            try
            {
                return doc.BlockRecords[name];
            }
            catch { }

            var br = new BlockRecord(name);

            // Put geometry on Layer 0 so it inherits the insert layer.
            Layer layer0 = null;
            try { layer0 = doc.Layers["0"]; } catch { }

            Line L(double x1, double y1, double x2, double y2)
            {
                var ln = new Line
                {
                    StartPoint = new XYZ(x1, y1, 0.0),
                    EndPoint = new XYZ(x2, y2, 0.0)
                };
                if (layer0 != null)
                    ln.Layer = layer0;
                return ln;
            }

            br.Entities.Add(L(0.0, 0.0, sheetWmm, 0.0));
            br.Entities.Add(L(sheetWmm, 0.0, sheetWmm, sheetHmm));
            br.Entities.Add(L(sheetWmm, sheetHmm, 0.0, sheetHmm));
            br.Entities.Add(L(0.0, sheetHmm, 0.0, 0.0));

            try
            {
                doc.BlockRecords.Add(br);
            }
            catch
            {
                // ignore (might already exist)
            }

            // Prefer returning the registered table entry.
            try { return doc.BlockRecords[name]; } catch { return br; }
        }

        private static void TrySetLayer(Entity entity, Layer layer)
        {
            if (entity == null || layer == null)
                return;

            try { entity.Layer = layer; } catch { }
        }

        private static void DrawSheetOutline(
            double originXmm,
            double originYmm,
            double sheetWmm,
            double sheetHmm,
            BlockRecord modelSpace,
            string materialLabel,
            int sheetIndex,
            NestingMode mode)
        {
            // Sheet outline as a block (requested):
            // - ModelSpace only gets one INSERT per sheet
            // - Easier selection/editing in CAD
            if (_sheetOutlineBlock != null)
            {
                var sheetIns = new Insert(_sheetOutlineBlock)
                {
                    InsertPoint = new XYZ(originXmm, originYmm, 0.0),
                    Rotation = 0.0,
                    XScale = 1.0,
                    YScale = 1.0,
                    ZScale = 1.0
                };

                TrySetLayer(sheetIns, _layerNestSheets);
                modelSpace.Entities.Add(sheetIns);
            }
            else
            {
                // Fallback (should not happen): draw 4 lines.
                var a = new Line { StartPoint = new XYZ(originXmm, originYmm, 0.0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm, 0.0) };
                var b = new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm, 0.0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0.0) };
                var c = new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0.0), EndPoint = new XYZ(originXmm, originYmm + sheetHmm, 0.0) };
                var d = new Line { StartPoint = new XYZ(originXmm, originYmm + sheetHmm, 0.0), EndPoint = new XYZ(originXmm, originYmm, 0.0) };
                TrySetLayer(a, _layerNestSheets);
                TrySetLayer(b, _layerNestSheets);
                TrySetLayer(c, _layerNestSheets);
                TrySetLayer(d, _layerNestSheets);
                modelSpace.Entities.Add(a);
                modelSpace.Entities.Add(b);
                modelSpace.Entities.Add(c);
                modelSpace.Entities.Add(d);
            }

            // Bottom info bar (separate layer)
            double barH = NEST_INFO_BAR_HEIGHT_MM;

            // NOTE: No bottom bar lines are drawn (text only under sheet).


            string title =
                $"Sheet {sheetIndex}" +
                (string.IsNullOrWhiteSpace(materialLabel) ? "" : $" | {materialLabel}") +
                $" | {mode}";

            var titleText = new MText
            {
                Value = title,
                InsertPoint = new XYZ(originXmm + 10.0, originYmm - barH + 10.0, 0.0),
                Height = 18.0
            };

            TrySetLayer(titleText, _layerNestBottomText);
            modelSpace.Entities.Add(titleText);
        }

        private static void AddFillLabels(BlockRecord modelSpace, List<SheetContourState> sheets, long usableW, long usableH, double sheetWmm, double sheetHmm)
        {
            // NOTE: UsedArea2Abs is in "2x area" units (shoelace sum). Sheet area must match.
            long usableArea2Abs = 2L * usableW * usableH;
            foreach (var s in sheets)
            {
                double fill = usableArea2Abs > 0 ? (double)s.UsedArea2Abs / usableArea2Abs * 100.0 : 0.0;

                var mt = new MText
                {
                    Value = $"Fill: {fill:0.0}%",
                    InsertPoint = new XYZ(s.OriginXmm + sheetWmm - 260.0, s.OriginYmm - NEST_INFO_BAR_HEIGHT_MM + 10.0, 0.0),
                    Height = 18.0
                };

                TrySetLayer(mt, _layerNestBottomText);
                modelSpace.Entities.Add(mt);
            }
        }

        private static void AddFillLabels(BlockRecord modelSpace, List<SheetRectState> sheets, long usableW, long usableH, double sheetWmm, double sheetHmm)
        {
            long usableArea2Abs = 2L * usableW * usableH;
            foreach (var s in sheets)
            {
                double fill = usableArea2Abs > 0 ? (double)s.UsedArea2Abs / usableArea2Abs * 100.0 : 0.0;

                var mt = new MText
                {
                    Value = $"Fill: {fill:0.0}%",
                    InsertPoint = new XYZ(s.OriginXmm + sheetWmm - 260.0, s.OriginYmm - NEST_INFO_BAR_HEIGHT_MM + 10.0, 0.0),
                    Height = 18.0
                };

                TrySetLayer(mt, _layerNestBottomText);
                modelSpace.Entities.Add(mt);
            }
        }

        private static void AddPlacedToDwg(
            BlockRecord modelSpace,
            PartDefinition part,
            SheetContourState sheet,
            double boundaryBufferMm,
            long insertXScaled,
            long insertYScaled,
            double rotRad)
        {
            double insXmm = sheet.OriginXmm + boundaryBufferMm + (double)insertXScaled / SCALE;
            double insYmm = sheet.OriginYmm + boundaryBufferMm + (double)insertYScaled / SCALE;

            var ins = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insXmm, insYmm, 0.0),
                Rotation = rotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            };

            TrySetLayer(ins, _layerNestParts);
            modelSpace.Entities.Add(ins);
        }

        private static void AppendNestLog(
            string logPath,
            string thicknessFile,
            string material,
            double sheetW,
            double sheetH,
            double thicknessMm,
            double gapMm,
            double marginMm,
            int sheets,
            int parts,
            string outDwg,
            NestingMode mode)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Nest run:");
                sb.AppendLine("  Thickness file: " + Path.GetFileName(thicknessFile));
                sb.AppendLine("  Material: " + material);
                sb.AppendLine($"  Mode: {mode}");
                sb.AppendLine($"  Sheet: {sheetW:0.###} x {sheetH:0.###} mm");
                sb.AppendLine($"  Thickness(mm): {thicknessMm:0.###}");
                sb.AppendLine($"  Gap(mm): {gapMm:0.###}  (auto >= thickness)");
                sb.AppendLine($"  Margin(mm): {marginMm:0.###}");
                sb.AppendLine($"  Sheets used: {sheets}");
                sb.AppendLine($"  Total parts: {parts}");
                sb.AppendLine("  Output: " + Path.GetFileName(outDwg));
                sb.AppendLine(new string('-', 70));

                File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        private static string MakeSafeFileToken(string s)
        {
            s = (s ?? "").Trim();
            if (s.Length == 0) return "UNKNOWN";

            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');

            s = s.Replace(' ', '_');

            if (s.Length > 80)
                s = s.Substring(0, 80);

            return s.Length == 0 ? "UNKNOWN" : s;
        }
    }
}
