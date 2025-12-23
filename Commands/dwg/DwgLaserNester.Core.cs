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
                int colMat = FindCol(headerFields, "Material", "SWMaterial", "SolidWorksMaterial");

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
            for (int i = 0; i < headerFields.Count; i++)
            {
                var h = (headerFields[i] ?? "").Trim().Trim('"').Trim();
                foreach (var n in names)
                {
                    if (string.Equals(h, n, StringComparison.OrdinalIgnoreCase))
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

                HashSet<string> matsFromCsv = null;
                if (idx != null && thickness > 0 && idx.MaterialsByThicknessKey.TryGetValue(thkKey, out var set) && set.Count > 0)
                    matsFromCsv = set;

                if (matsFromCsv != null)
                {
                    foreach (var mat in matsFromCsv.OrderBy(s => s, StringComparer.Ordinal))
                    {
                        jobs.Add(new LaserNestJob
                        {
                            Enabled = true,
                            ThicknessFilePath = file,
                            ThicknessMm = thickness,
                            MaterialExact = mat
                        });
                    }
                    continue;
                }

                // Fallback: unknown
                jobs.Add(new LaserNestJob
                {
                    Enabled = true,
                    ThicknessFilePath = file,
                    ThicknessMm = thickness,
                    MaterialExact = "UNKNOWN"
                });
            }

            return jobs
                .OrderBy(j => j.MaterialExact ?? "", StringComparer.Ordinal)
                .ThenBy(j => j.ThicknessMm <= 0 ? double.MaxValue : j.ThicknessMm)
                .ThenBy(j => j.ThicknessFileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();
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

            var summary = new StringBuilder();
            summary.AppendLine("Batch nesting summary");
            summary.AppendLine("Folder: " + folder);
            summary.AppendLine("Tasks: " + selected.Count);
            summary.AppendLine("Mode: " + settings.Mode);
            summary.AppendLine("Note: SeparateByMaterialExact / OneDWGPerMaterial / FilterPreview are FORCED true.");
            summary.AppendLine(new string('-', 70));

            string summaryPath = Path.Combine(folder, "batch_nest_summary.txt");

            bool cancelled = false;

            using (var progress = new LaserCutProgressForm())
            {
                progress.Show();
                progress.BeginBatch(selected.Count);

                int done = 0;

                try
                {
                    foreach (var job in selected)
                    {
                        progress.ThrowIfCancelled();

                        done++;

                        var result = RunSingleJob(job, settings, progress, done, selected.Count);

                        summary.AppendLine(Path.GetFileName(result.ThicknessFile));
                        summary.AppendLine($"  Material: {result.MaterialExact}");
                        summary.AppendLine($"  Mode: {result.Mode}");
                        summary.AppendLine($"  SheetsUsed: {result.SheetsUsed}, Parts: {result.TotalParts}");
                        summary.AppendLine($"  Output: {Path.GetFileName(result.OutputDwg)}");
                        summary.AppendLine(new string('-', 70));

                        progress.EndTask(done);
                    }
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                    summary.AppendLine("CANCELLED BY USER");
                    summary.AppendLine(new string('-', 70));
                }
                catch (Exception ex)
                {
                    summary.AppendLine("ERROR:");
                    summary.AppendLine(ex.ToString());
                    summary.AppendLine(new string('-', 70));
                    throw;
                }
                finally
                {
                    try { progress.Close(); } catch { }
                }
            }

            File.WriteAllText(summaryPath, summary.ToString(), Encoding.UTF8);

            if (showUi)
            {
                MessageBox.Show(
                    cancelled
                        ? "Nesting was cancelled.\r\n\r\nSummary:\r\n" + summaryPath
                        : "Batch nesting finished.\r\n\r\nSummary:\r\n" + summaryPath,
                    "Laser nesting",
                    MessageBoxButtons.OK,
                    cancelled ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }

        // ============================
        // Single job
        // ============================
        private static NestRunResult RunSingleJob(LaserNestJob job, LaserCutRunSettings settings, LaserCutProgressForm progress, int taskIndex, int totalTasks)
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

            int totalInstances = defs.Sum(d => d.Quantity);

            // if mapping failed, fallback to all (still produce output)
            if (totalInstances <= 0)
            {
                defs = defsAll;
                totalInstances = defs.Sum(d => d.Quantity);
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

            progress.BeginTask(
                taskIndex,
                totalTasks,
                thicknessFileName,
                material,
                thicknessMm,
                totalInstances,
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
                        totalInstances);
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
                        totalInstances,
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
                        totalInstances,
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
                    thicknessFileName,
                    material,
                    thicknessMm,
                    totalInstances,
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
                    totalInstances,
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
                totalInstances,
                outPath,
                finalMode);

            return new NestRunResult
            {
                ThicknessFile = job.ThicknessFilePath,
                MaterialExact = material,
                OutputDwg = outPath,
                SheetsUsed = sheetsUsed,
                TotalParts = totalInstances,
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
