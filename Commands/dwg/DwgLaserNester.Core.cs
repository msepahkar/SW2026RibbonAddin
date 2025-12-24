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


            // Optional optimization: mirror-pair parts that together fill a rectangle
            // (common-line). This is OFF by default (see UI checkbox).
            if (settings.MirrorPairing != MirrorPairingMode.Off)
            {
                ApplyMirrorPairingToRectangles(doc, defs, material, settings.MirrorPairing, gapMm);
                totalParts = defs.Sum(d => d.Quantity * Math.Max(1, d.PartCountWeight));
            }

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


        // ============================
        // Mirror-pair optimization
        private enum MirrorPairAxis { X = 0, Y = 1 }

        // ============================
        private sealed class _MirrorShapeInfo
        {
            public PartDefinition Part;
            // Rotation applied to the original block (radians) so that the
            // resulting pair-rectangle becomes axis-aligned.
            public double RotRad;

            // Rotated polygon, normalized so its bbox min is at (0,0).
            // Full precision version is used for geometric verification.
            public Path64 PolyNormFull;

            // Decimated + snapped version used for hashing/matching.
            public Path64 PolyNormDec;

            // Bounding box of the rotated polygon BEFORE normalization.
            // Used to compute the insertion translation for the rotated block.
            public LongRect RotBounds0;
            public long Width;
            public long Height;
            public ulong Key;
            public ulong MirrorXKey;
            public ulong MirrorYKey;
        }

        /// <summary>
        /// Attempts to detect mirrored part pairs that, when overlaid in the same bounding box,
        /// form a full rectangle (no overlap area, union fills the rectangle).
        ///
        /// If found, it creates a new "PAIR_*" block (containing two inserts) and adds a
        /// synthetic PartDefinition with PartCountWeight=2 and a rectangular contour.
        ///
        /// This is especially helpful for Fast(Rectangles) mode when individual parts have
        /// large bounding-box waste (L shapes, triangles, etc.).
        /// </summary>
        private static void ApplyMirrorPairingToRectangles(CadDocument doc, List<PartDefinition> defs, string materialLabel, MirrorPairingMode mode, double pairGapMm)
        {
            if (doc == null || defs == null || defs.Count < 2)
                return;

            // Only consider parts with positive qty.
            var parts = defs.Where(d => d != null && d.Quantity > 0).ToList();
            if (parts.Count < 2)
                return;

            const int HASH_MAX_POINTS = 220;
            const double HASH_SNAP_MM = 0.01; // helps stable hashing after rotation
            const int MAX_ANGLES_PER_PART = 14;

            // Precompute rotated-normalized polys + hash keys for a small set of
            // candidate angles per part. This enables mirror-pair detection even
            // when the parts are exported in arbitrary orientations.
            var infosByPart = new Dictionary<PartDefinition, List<_MirrorShapeInfo>>();
            var byKey = new Dictionary<ulong, List<_MirrorShapeInfo>>();

            foreach (var p in parts)
            {
                Path64 poly0 = p.OuterContour0;
                if (poly0 == null || poly0.Count < 3)
                    poly0 = MakeRectPolyScaled(p.MinX, p.MinY, p.MaxX, p.MaxY);

                poly0 = CleanPath(poly0);
                if (poly0 == null || poly0.Count < 3)
                    continue;

                var angles = GetCandidateAnglesForMirrorPairing(poly0, MAX_ANGLES_PER_PART);
                if (angles == null || angles.Count == 0)
                    angles = new List<double> { 0.0 };

                var list = new List<_MirrorShapeInfo>();

                foreach (double a in angles)
                {
                    // Rotate by -a so that an edge at angle 'a' becomes horizontal.
                    double rotRad = -a;

                    var polyRot = RotatePolyRad(poly0, rotRad);
                    polyRot = CleanPath(polyRot);
                    if (polyRot == null || polyRot.Count < 3)
                        continue;

                    var bRot = GetBounds(polyRot);
                    long w = bRot.MaxX - bRot.MinX;
                    long h = bRot.MaxY - bRot.MinY;
                    if (w <= 0 || h <= 0)
                        continue;

                    var normFull = TranslatePath(polyRot, -bRot.MinX, -bRot.MinY);
                    normFull = CleanPath(normFull);

                    var normDec = CleanPath(DecimatePath(normFull, HASH_MAX_POINTS));
                    normDec = SnapPath(normDec, HASH_SNAP_MM);

                    if (normDec == null || normDec.Count < 3)
                        continue;

                    // Ensure snapped decimated shape is also anchored at 0,0.
                    var bDec = GetBounds(normDec);
                    if (bDec.MinX != 0 || bDec.MinY != 0)
                        normDec = CleanPath(TranslatePath(normDec, -bDec.MinX, -bDec.MinY));

                    if (normDec == null || normDec.Count < 3)
                        continue;

                    ulong key = HashPathCanonical(normDec);
                    if (key == 0UL)
                        continue;

                    var bHash = GetBounds(normDec);
                    long wHash = bHash.MaxX - bHash.MinX;
                    long hHash = bHash.MaxY - bHash.MinY;
                    if (wHash <= 0 || hHash <= 0)
                        continue;

                    var mx = MirrorX(normDec, wHash);
                    var my = MirrorY(normDec, hHash);
                    ulong kx = HashPathCanonical(mx);
                    ulong ky = HashPathCanonical(my);

                    var info = new _MirrorShapeInfo
                    {
                        Part = p,
                        RotRad = rotRad,
                        PolyNormFull = normFull,
                        PolyNormDec = normDec,
                        RotBounds0 = bRot,
                        Width = w,
                        Height = h,
                        Key = key,
                        MirrorXKey = kx,
                        MirrorYKey = ky,
                    };

                    list.Add(info);

                    if (!byKey.TryGetValue(key, out var klist))
                    {
                        klist = new List<_MirrorShapeInfo>();
                        byKey[key] = klist;
                    }
                    klist.Add(info);
                }

                if (list.Count > 0)
                    infosByPart[p] = list;
            }

            if (infosByPart.Count < 2)
                return;

            // Process bigger parts first (more benefit).
            var orderedParts = infosByPart.Keys
                .OrderByDescending(d => Math.Max(0.0, d.Width) * Math.Max(0.0, d.Height))
                .ToList();

            int createdPairs = 0;

            foreach (var A in orderedParts)
            {
                if (A == null || A.Quantity <= 0)
                    continue;

                if (!infosByPart.TryGetValue(A, out var aInfos) || aInfos == null || aInfos.Count == 0)
                    continue;

                // Try to consume as many mirror pairs for this part as possible.
                // Usually there is only one matching mirror-part, but this loop
                // allows multiple matches if the data contains duplicates.
                bool pairedSomething;
                do
                {
                    pairedSomething = false;

                    if (A.Quantity <= 0)
                        break;

                    _MirrorShapeInfo bestA = null;
                    _MirrorShapeInfo bestB = null;
                    MirrorPairAxis bestAxis = MirrorPairAxis.X;

                    // Find the first valid candidate (verified by geometry).
                    foreach (var a in aInfos)
                    {
                        if (A.Quantity <= 0)
                            break;

                        if (a == null || a.Key == 0UL)
                            continue;

                        if (TryFindVerifiedMirrorCandidate(a, byKey, out var b, out var axis))
                        {
                            bestA = a;
                            bestB = b;
                            bestAxis = axis;
                            break;
                        }
                    }

                    if (bestA == null || bestB == null || bestB.Part == null)
                        break;

                    int pairCount = Math.Min(A.Quantity, bestB.Part.Quantity);
                    if (pairCount <= 0)
                        break;

                    if (CreateMirrorPairBlock(doc, defs, materialLabel, bestA, bestB, bestAxis, mode, pairGapMm, pairCount, ref createdPairs))
                    {
                        // Consume quantities
                        A.Quantity -= pairCount;
                        bestB.Part.Quantity -= pairCount;
                        pairedSomething = true;
                    }

                } while (pairedSomething);
            }

            // Remove zero-qty originals (keeps list smaller for placement)
            defs.RemoveAll(d => d == null || d.Quantity <= 0);
        }

        private static bool TryFindVerifiedMirrorCandidate(
            _MirrorShapeInfo a,
            Dictionary<ulong, List<_MirrorShapeInfo>> byKey,
            out _MirrorShapeInfo bBest,
            out MirrorPairAxis axisBest)
        {
            bBest = null;
            axisBest = MirrorPairAxis.X;

            if (a == null || a.Part == null || a.Part.Quantity <= 0)
                return false;

            // Try mirror-X, then mirror-Y.
            if (TryFindVerifiedMirrorCandidateAxis(a, byKey, MirrorPairAxis.X, out bBest))
            {
                axisBest = MirrorPairAxis.X;
                return true;
            }

            if (TryFindVerifiedMirrorCandidateAxis(a, byKey, MirrorPairAxis.Y, out bBest))
            {
                axisBest = MirrorPairAxis.Y;
                return true;
            }

            return false;
        }

        private static bool TryFindVerifiedMirrorCandidateAxis(
            _MirrorShapeInfo a,
            Dictionary<ulong, List<_MirrorShapeInfo>> byKey,
            MirrorPairAxis axis,
            out _MirrorShapeInfo bBest)
        {
            bBest = null;

            ulong mirrorKey = axis == MirrorPairAxis.X ? a.MirrorXKey : a.MirrorYKey;
            if (mirrorKey == 0UL)
                return false;

            if (!byKey.TryGetValue(mirrorKey, out var candidates) || candidates == null || candidates.Count == 0)
                return false;

            foreach (var b in candidates)
            {
                var B = b?.Part;
                if (B == null)
                    continue;

                if (ReferenceEquals(B, a.Part))
                    continue;

                if (a.Part.Quantity <= 0)
                    break;

                if (B.Quantity <= 0)
                    continue;

                // Must match bounding box dims very closely.
                if (!DimsClose(a.Width, b.Width) || !DimsClose(a.Height, b.Height))
                    continue;

                // Expensive verify: union fills the rectangle and intersection area is ~0.
                if (!VerifyRectangleUnionNoOverlap(a.PolyNormFull, b.PolyNormFull, a.Width, a.Height))
                    continue;

                bBest = b;
                return true;
            }

            return false;
        }

        private static bool CreateMirrorPairBlock(
            CadDocument doc,
            List<PartDefinition> defs,
            string materialLabel,
            _MirrorShapeInfo a,
            _MirrorShapeInfo b,
            MirrorPairAxis axis,
            MirrorPairingMode mode,
            double pairGapMm,
            int pairCount,
            ref int createdPairs)
        {
            if (doc == null || defs == null || a == null || b == null || a.Part == null || b.Part == null)
                return false;

            var A = a.Part;
            var B = b.Part;

            // Create a new pair block definition once per A-B match.
            string pairBlockName = BuildPairBlockName(A.BlockName, B.BlockName, createdPairs + 1);

            BlockRecord pairBlock = null;
            string nameTry = pairBlockName;

            for (int attempt = 0; attempt < 200; attempt++)
            {
                try
                {
                    pairBlock = new BlockRecord(nameTry);
                    doc.BlockRecords.Add(pairBlock);
                    break;
                }
                catch
                {
                    nameTry = pairBlockName + "_" + (attempt + 1).ToString(CultureInfo.InvariantCulture);
                    pairBlock = null;
                }
            }

            if (pairBlock == null)
                return false;

            // CommonLine: the two parts touch (common-line) and together form an exact rectangle.
            // WithGap: keep an internal gap (use auto gap) by shifting one half, and grow the outer rectangle by that gap.
            long gapScaled = 0;
            if (mode == MirrorPairingMode.WithGap && pairGapMm > 0)
                gapScaled = (long)Math.Round(pairGapMm * SCALE);

            long rectW = a.Width;
            long rectH = a.Height;

            double extraBx = 0.0;
            double extraBy = 0.0;

            if (gapScaled > 0)
            {
                if (axis == MirrorPairAxis.X)
                {
                    rectW = a.Width + gapScaled;
                    extraBx = (double)gapScaled / SCALE;
                }
                else if (axis == MirrorPairAxis.Y)
                {
                    rectH = a.Height + gapScaled;
                    extraBy = (double)gapScaled / SCALE;
                }
            }

            // Insert translations: move each rotated block so its rotated bbox min sits at (0,0).
            double ax = -(double)a.RotBounds0.MinX / SCALE;
            double ay = -(double)a.RotBounds0.MinY / SCALE;
            double bx = -(double)b.RotBounds0.MinX / SCALE + extraBx;
            double by = -(double)b.RotBounds0.MinY / SCALE + extraBy;

            pairBlock.Entities.Add(new Insert(A.Block)
            {
                InsertPoint = new XYZ(ax, ay, 0.0),
                Rotation = a.RotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            });

            pairBlock.Entities.Add(new Insert(B.Block)
            {
                InsertPoint = new XYZ(bx, by, 0.0),
                Rotation = b.RotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            });

            // New synthetic part def representing a rectangle that yields TWO parts.
            double wMm = (double)rectW / SCALE;
            double hMm = (double)rectH / SCALE;

            var rect = new Path64
            {
                new Point64(0, 0),
                new Point64(rectW, 0),
                new Point64(rectW, rectH),
                new Point64(0, rectH)
            };

            defs.Add(new PartDefinition
            {
                Block = pairBlock,
                BlockName = pairBlock.Name,
                Quantity = pairCount,
                PartCountWeight = 2,
                MaterialExact = materialLabel ?? "UNKNOWN",

                MinX = 0.0,
                MinY = 0.0,
                MaxX = wMm,
                MaxY = hMm,
                Width = wMm,
                Height = hMm,

                OuterContour0 = rect,
                OuterArea2Abs = Area2Abs(rect)
            });

            createdPairs++;
            return true;
        }

        private static List<double> GetCandidateAnglesForMirrorPairing(Path64 poly, int maxAngles)
        {
            if (poly == null || poly.Count < 3)
                return new List<double> { 0.0 };

            // Reduce point count a bit to avoid noisy micro-edges (especially from arcs).
            var p = CleanPath(DecimatePath(poly, 520));
            if (p == null || p.Count < 3)
                return new List<double> { 0.0 };

            // Collect edge angles weighted by edge length.
            var cands = new List<(double ang, double w)>();
            int n = p.Count;

            // Ignore very small edges (< ~2mm) to avoid noise.
            double minLen = 2.0 * SCALE;
            double minLen2 = minLen * minLen;

            for (int i = 0; i < n; i++)
            {
                var a = p[i];
                var b = p[(i + 1) % n];
                long dxL = b.X - a.X;
                long dyL = b.Y - a.Y;
                double dx = dxL;
                double dy = dyL;
                double len2 = dx * dx + dy * dy;
                if (len2 < minLen2)
                    continue;

                double ang = Math.Atan2(dy, dx);
                double norm = NormalizeAngle0ToHalfPi(ang);
                cands.Add((norm, len2));
            }

            // Always include 0.
            var result = new List<double> { 0.0 };

            if (cands.Count == 0)
                return result;

            // Sort longest edges first.
            cands.Sort((x, y) => y.w.CompareTo(x.w));

            double tol = 0.5 * Math.PI / 180.0; // 0.5 degree

            foreach (var c in cands)
            {
                if (result.Count >= Math.Max(2, maxAngles))
                    break;

                bool close = false;
                foreach (double existing in result)
                {
                    double diff = Math.Abs(c.ang - existing);
                    diff = Math.Min(diff, (Math.PI / 2.0) - diff);
                    if (diff < tol)
                    {
                        close = true;
                        break;
                    }
                }

                if (!close)
                    result.Add(c.ang);
            }

            return result;
        }

        private static double NormalizeAngle0ToHalfPi(double angleRad)
        {
            // Map to [0, PI)
            double a = angleRad % Math.PI;
            if (a < 0) a += Math.PI;

            // Fold to [0, PI/2)
            if (a >= Math.PI / 2.0)
                a -= Math.PI / 2.0;

            return a;
        }

        private static bool DimsClose(long a, long b)
        {
            // within 0.05mm (snap tolerance default) in scaled units
            long tol = Math.Max(1, (long)Math.Round(0.05 * SCALE));
            return Math.Abs(a - b) <= tol;
        }

        private static bool VerifyRectangleUnionNoOverlap(Path64 aNorm, Path64 bNorm, long w, long h)
        {
            if (aNorm == null || bNorm == null || aNorm.Count < 3 || bNorm.Count < 3)
                return false;

            if (w <= 0 || h <= 0)
                return false;

            // Intersection area should be ~0
            var clip = new Clipper64();
            clip.AddSubject(aNorm);
            clip.AddClip(bNorm);

            var inter = new Paths64();
            clip.Execute(Clipper2Lib.ClipType.Intersection, FillRule.NonZero, inter);

            long interA2 = 0;
            if (inter != null)
            {
                foreach (var p in inter)
                    interA2 += Area2Abs(p);
            }

            // Allow tiny numerical noise (≈ 2mm^2)
            const long interTolA2 = 4_000_000; // 2mm^2 -> 2*1e6
            if (interA2 > interTolA2)
                return false;

            // Union area should be ~= rectangle area
            var clipU = new Clipper64();
            clipU.AddSubject(aNorm);
            clipU.AddSubject(bNorm);

            var uni = new Paths64();
            clipU.Execute(Clipper2Lib.ClipType.Union, FillRule.NonZero, uni);

            long uniA2 = 0;
            if (uni != null)
            {
                foreach (var p in uni)
                    uniA2 += Area2Abs(p);
            }

            long rectA2 = 2L * w * h;

            // Relative tolerance 0.5% + small absolute tolerance
            long absTol = 6_000_000; // 3mm^2
            long relTol = (long)Math.Round(rectA2 * 0.005);

            long tol = Math.Max(absTol, relTol);

            return Math.Abs(uniA2 - rectA2) <= tol;
        }

        private static Path64 MirrorX(Path64 p, long width)
        {
            if (p == null) return p;

            var r = new Path64(p.Count);
            foreach (var pt in p)
                r.Add(new Point64(width - pt.X, pt.Y));

            return CleanPath(r);
        }

        private static Path64 MirrorY(Path64 p, long height)
        {
            if (p == null) return p;

            var r = new Path64(p.Count);
            foreach (var pt in p)
                r.Add(new Point64(pt.X, height - pt.Y));

            return CleanPath(r);
        }

        private static string BuildPairBlockName(string aName, string bName, int index)
        {
            string Clean(string s, int maxLen)
            {
                s = (s ?? "").Trim();
                if (s.Length == 0) return "X";

                var sb = new StringBuilder(s.Length);
                foreach (char c in s)
                {
                    if (char.IsLetterOrDigit(c))
                        sb.Append(c);
                    else if (c == '_' || c == '-')
                        sb.Append('_');
                }

                var t = sb.ToString();
                if (t.Length == 0) t = "X";
                if (t.Length > maxLen) t = t.Substring(0, maxLen);
                return t;
            }

            string aa = Clean(aName, 24);
            string bb = Clean(bName, 24);

            return $"PAIR_{index}_{aa}_{bb}";
        }

        // Canonical hash for a closed polygon path, invariant to start index and direction.
        private static ulong HashPathCanonical(Path64 p)
        {
            if (p == null || p.Count < 3)
                return 0UL;

            p = CleanPath(p);
            int n = p.Count;
            if (n < 3)
                return 0UL;

            long minX = long.MaxValue;
            long minY = long.MaxValue;

            for (int i = 0; i < n; i++)
            {
                var pt = p[i];
                if (pt.X < minX || (pt.X == minX && pt.Y < minY))
                {
                    minX = pt.X;
                    minY = pt.Y;
                }
            }

            var starts = new List<int>();
            for (int i = 0; i < n; i++)
                if (p[i].X == minX && p[i].Y == minY)
                    starts.Add(i);

            int bestStart = starts[0];
            int bestDir = +1; // +1 forward, -1 reverse

            foreach (int s in starts)
            {
                if (CompareCyclic(p, s, +1, bestStart, bestDir) < 0)
                {
                    bestStart = s;
                    bestDir = +1;
                }

                if (CompareCyclic(p, s, -1, bestStart, bestDir) < 0)
                {
                    bestStart = s;
                    bestDir = -1;
                }
            }

            // FNV-1a 64-bit
            ulong h = 1469598103934665603UL;
            const ulong prime = 1099511628211UL;

            for (int k = 0; k < n; k++)
            {
                int idx = bestDir > 0 ? (bestStart + k) % n : (bestStart - k + n) % n;
                var pt = p[idx];

                unchecked
                {
                    h ^= (ulong)pt.X;
                    h *= prime;
                    h ^= (ulong)pt.Y;
                    h *= prime;
                }
            }

            return h;
        }

        private static int CompareCyclic(Path64 p, int startA, int dirA, int startB, int dirB)
        {
            int n = p.Count;

            for (int k = 0; k < n; k++)
            {
                int ia = dirA > 0 ? (startA + k) % n : (startA - k + n) % n;
                int ib = dirB > 0 ? (startB + k) % n : (startB - k + n) % n;

                var a = p[ia];
                var b = p[ib];

                if (a.X != b.X)
                    return a.X < b.X ? -1 : 1;
                if (a.Y != b.Y)
                    return a.Y < b.Y ? -1 : 1;
            }

            return 0;
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
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm, originYmm, 0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm, 0), EndPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm + sheetWmm, originYmm + sheetHmm, 0), EndPoint = new XYZ(originXmm, originYmm + sheetHmm, 0) });
            modelSpace.Entities.Add(new Line { StartPoint = new XYZ(originXmm, originYmm + sheetHmm, 0), EndPoint = new XYZ(originXmm, originYmm, 0) });

            string title =
                $"Sheet {sheetIndex}" +
                (string.IsNullOrWhiteSpace(materialLabel) ? "" : $" | {materialLabel}") +
                $" | {mode}";

            modelSpace.Entities.Add(new MText
            {
                Value = title,
                InsertPoint = new XYZ(originXmm + 10.0, originYmm + sheetHmm + 18.0, 0.0),
                Height = 20.0
            });
        }

        private static void AddFillLabels(BlockRecord modelSpace, List<SheetContourState> sheets, long usableW, long usableH, double sheetWmm, double sheetHmm)
        {
            long usableArea2 = usableW * usableH;
            foreach (var s in sheets)
            {
                double fill = usableArea2 > 0 ? (double)s.UsedArea2Abs / usableArea2 * 100.0 : 0.0;

                modelSpace.Entities.Add(new MText
                {
                    Value = $"Fill: {fill:0.0}%",
                    InsertPoint = new XYZ(s.OriginXmm + sheetWmm - 220.0, s.OriginYmm + sheetHmm + 18.0, 0.0),
                    Height = 18.0
                });
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
