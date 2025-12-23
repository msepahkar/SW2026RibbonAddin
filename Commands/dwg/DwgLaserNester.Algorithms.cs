using ACadSharp.Entities;
using ACadSharp.Tables;
using Clipper2Lib;
using CSMath;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace SW2026RibbonAddin.Commands
{
    internal static partial class DwgLaserNester
    {
        // ============================
        // Level 2 production settings
        // ============================
        private const int LEVEL2_JOB_TIMEOUT_SECONDS = 120; // hard safety: job-level time budget
        private const int LEVEL2_MAX_CANDIDATES_CAP = 800;
        private const int LEVEL2_MAX_PARTNERS_CAP = 20;
        private const int LEVEL2_NFP_MAX_POINTS = 180;

        // HYBRID: use Level2 for “big parts”, then Level1 for the rest (fast & safe)
        private const int LEVEL2_HYBRID_MIN_TOPN = 10;
        private const int LEVEL2_HYBRID_MAX_TOPN = 60;
        private const int LEVEL2_HYBRID_DIVISOR = 5; // total/5 => 20%

        private static int ComputeHybridTopN(int totalInstances)
        {
            totalInstances = Math.Max(0, totalInstances);
            if (totalInstances <= 0) return 0;

            int byFraction = totalInstances / LEVEL2_HYBRID_DIVISOR;
            int top = Math.Max(LEVEL2_HYBRID_MIN_TOPN, byFraction);
            top = Math.Min(top, LEVEL2_HYBRID_MAX_TOPN);
            top = Math.Min(top, totalInstances);
            return top;
        }

        internal sealed class Level2TimeoutException : Exception
        {
            public Level2TimeoutException(string msg) : base(msg) { }
        }

        private static long MakeDeadlineTicks(int seconds)
        {
            long now = Stopwatch.GetTimestamp();
            long add = (long)(seconds * (double)Stopwatch.Frequency);
            return now + Math.Max(1, add);
        }

        private static void ThrowIfDeadlineExceeded(long deadlineTicks)
        {
            if (Stopwatch.GetTimestamp() > deadlineTicks)
                throw new Level2TimeoutException("Level 2 timeout exceeded.");
        }

        // ---------------------------
        // FAST MODE (Rectangles)
        // ---------------------------
        private static int NestFastRectangles(
            List<PartDefinition> defs,
            BlockRecord modelSpace,
            double sheetWmm,
            double sheetHmm,
            double marginMm,
            double gapMm,
            double baseOriginXmm,
            double baseOriginYmm,
            string materialLabel,
            LaserCutProgressForm progress,
            int totalInstances)
        {
            progress.ThrowIfCancelled();

            double placementMargin = marginMm + gapMm;

            double usableW = sheetWmm - 2 * placementMargin;
            double usableH = sheetHmm - 2 * placementMargin;
            if (usableW <= 0 || usableH <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            var instances = ExpandInstances(defs);
            instances.Sort((a, b) => (b.Width * b.Height).CompareTo(a.Width * a.Height));

            var sheets = new List<SheetRectState>();

            SheetRectState NewSheet()
            {
                var s = new SheetRectState
                {
                    Index = sheets.Count + 1,
                    OriginXmm = baseOriginXmm + (sheets.Count) * (sheetWmm + 60.0),
                    OriginYmm = baseOriginYmm
                };

                DrawSheetOutline(s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, modelSpace, materialLabel, s.Index, NestingMode.FastRectangles);

                s.Free.Add(new FreeRect
                {
                    X = placementMargin,
                    Y = placementMargin,
                    W = sheetWmm - 2 * placementMargin,
                    H = sheetHmm - 2 * placementMargin
                });

                sheets.Add(s);
                return s;
            }

            var cur = NewSheet();

            int placed = 0;
            progress.ReportPlaced(0, totalInstances, sheets.Count);

            foreach (var part in instances)
            {
                progress.ThrowIfCancelled();

                while (true)
                {
                    if (TryPlaceRect(cur, part, gapMm, modelSpace))
                    {
                        placed++;
                        progress.ReportPlaced(placed, totalInstances, sheets.Count);
                        break;
                    }

                    cur = NewSheet();
                    progress.ReportPlaced(placed, totalInstances, sheets.Count);
                }
            }

            return sheets.Count;
        }

        private static List<PartDefinition> ExpandInstances(List<PartDefinition> defs)
        {
            var list = new List<PartDefinition>();
            foreach (var d in defs)
                for (int i = 0; i < d.Quantity; i++)
                    list.Add(d);
            return list;
        }

        private static bool TryPlaceRect(SheetRectState sheet, PartDefinition part, double gapMm, BlockRecord modelSpace)
        {
            for (int frIndex = 0; frIndex < sheet.Free.Count; frIndex++)
            {
                var fr = sheet.Free[frIndex];

                if (TryPlaceRectOrientation(sheet, part, frIndex, fr, part.Width, part.Height, 0.0, gapMm, modelSpace))
                    return true;

                if (TryPlaceRectOrientation(sheet, part, frIndex, fr, part.Height, part.Width, Math.PI / 2.0, gapMm, modelSpace))
                    return true;
            }

            return false;
        }

        private static bool TryPlaceRectOrientation(
            SheetRectState sheet,
            PartDefinition part,
            int frIndex,
            FreeRect fr,
            double w,
            double h,
            double rotRad,
            double gapMm,
            BlockRecord modelSpace)
        {
            double usedW = w + gapMm;
            double usedH = h + gapMm;

            if (usedW > fr.W || usedH > fr.H)
                return false;

            double localMinX = fr.X + gapMm * 0.5;
            double localMinY = fr.Y + gapMm * 0.5;

            double worldMinX = sheet.OriginXmm + localMinX;
            double worldMinY = sheet.OriginYmm + localMinY;

            double insertX;
            double insertY;

            if (Math.Abs(rotRad) < 1e-9)
            {
                insertX = worldMinX - part.MinX;
                insertY = worldMinY - part.MinY;
            }
            else
            {
                insertX = worldMinX + part.MaxY;
                insertY = worldMinY - part.MinX;
            }

            var ins = new Insert(part.Block)
            {
                InsertPoint = new XYZ(insertX, insertY, 0.0),
                Rotation = rotRad,
                XScale = 1.0,
                YScale = 1.0,
                ZScale = 1.0
            };
            modelSpace.Entities.Add(ins);

            sheet.Free.RemoveAt(frIndex);

            double rightW = fr.W - usedW;
            if (rightW > 1.0)
                sheet.Free.Add(new FreeRect { X = fr.X + usedW, Y = fr.Y, W = rightW, H = fr.H });

            double topH = fr.H - usedH;
            if (topH > 1.0)
                sheet.Free.Add(new FreeRect { X = fr.X, Y = fr.Y + usedH, W = usedW, H = topH });

            return true;
        }

        // ---------------------------
        // CONTOUR LEVEL 1
        // ---------------------------
        private static int NestContourLevel1(
            List<PartDefinition> defs,
            BlockRecord modelSpace,
            double sheetWmm,
            double sheetHmm,
            double marginMm,
            double gapMm,
            double baseOriginXmm,
            double baseOriginYmm,
            string materialLabel,
            LaserCutProgressForm progress,
            int totalInstances,
            double chordMm,
            double snapMm,
            int maxCandidates)
        {
            progress.ThrowIfCancelled();

            double boundaryBufferMm = marginMm + gapMm / 2.0;

            double usableWmm = sheetWmm - 2 * boundaryBufferMm;
            double usableHmm = sheetHmm - 2 * boundaryBufferMm;
            if (usableWmm <= 0 || usableHmm <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            long usableW = ToInt(usableWmm);
            long usableH = ToInt(usableHmm);

            var instances = ExpandInstances(defs);
            instances.Sort((a, b) => SortByAreaDesc(a, b));

            var polyCache = new Dictionary<string, RotatedPoly>(StringComparer.OrdinalIgnoreCase);
            RotatedPoly GetRot(PartDefinition part, int rotDeg) => GetOrCreateRotated(part, rotDeg, gapMm, polyCache);

            var sheets = new List<SheetContourState>();

            SheetContourState NewSheet()
            {
                var s = new SheetContourState
                {
                    Index = sheets.Count + 1,
                    OriginXmm = baseOriginXmm + (sheets.Count) * (sheetWmm + 60.0),
                    OriginYmm = baseOriginYmm
                };

                DrawSheetOutline(s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, modelSpace, materialLabel, s.Index, NestingMode.ContourLevel1);
                sheets.Add(s);
                return s;
            }

            var cur = NewSheet();

            int placed = 0;
            progress.ReportPlaced(0, totalInstances, sheets.Count);

            foreach (var part in instances)
            {
                progress.ThrowIfCancelled();

                bool placedThis = false;

                for (int si = 0; si < sheets.Count; si++)
                {
                    if (TryPlaceContourOnSheet_Level1(sheets[si], part, usableW, usableH, maxCandidates, GetRot, progress, out var placement))
                    {
                        AddPlacedToDwg(modelSpace, part, sheets[si], boundaryBufferMm, placement.InsertX, placement.InsertY, placement.RotRad);

                        var decNfp = CleanPath(DecimatePath(placement.OffsetPolyTranslated, LEVEL2_NFP_MAX_POINTS));

                        sheets[si].Placed.Add(new PlacedContour
                        {
                            OffsetPoly = placement.OffsetPolyTranslated,
                            OffsetPolyNfp = decNfp,
                            BBox = placement.OffsetBBoxTranslated
                        });

                        sheets[si].PlacedCount++;
                        sheets[si].UsedArea2Abs += placement.RotArea2Abs;

                        placed++;
                        progress.ReportPlaced(placed, totalInstances, sheets.Count);
                        placedThis = true;
                        break;
                    }
                }

                if (placedThis)
                    continue;

                cur = NewSheet();
                progress.ReportPlaced(placed, totalInstances, sheets.Count);

                if (!TryPlaceContourOnSheet_Level1(cur, part, usableW, usableH, maxCandidates, GetRot, progress, out var placement2))
                    throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

                AddPlacedToDwg(modelSpace, part, cur, boundaryBufferMm, placement2.InsertX, placement2.InsertY, placement2.RotRad);

                var decNfp2 = CleanPath(DecimatePath(placement2.OffsetPolyTranslated, LEVEL2_NFP_MAX_POINTS));

                cur.Placed.Add(new PlacedContour
                {
                    OffsetPoly = placement2.OffsetPolyTranslated,
                    OffsetPolyNfp = decNfp2,
                    BBox = placement2.OffsetBBoxTranslated
                });

                cur.PlacedCount++;
                cur.UsedArea2Abs += placement2.RotArea2Abs;

                placed++;
                progress.ReportPlaced(placed, totalInstances, sheets.Count);
            }

            AddFillLabels(modelSpace, sheets, usableW, usableH, sheetWmm, sheetHmm);
            return sheets.Count;
        }

        private static bool TryPlaceContourOnSheet_Level1(
            SheetContourState sheet,
            PartDefinition part,
            long usableW,
            long usableH,
            int maxCandidates,
            Func<PartDefinition, int, RotatedPoly> getRot,
            LaserCutProgressForm progress,
            out ContourPlacement placement)
        {
            placement = default;

            foreach (int rotDeg in RotationsDeg)
            {
                progress.ThrowIfCancelled();

                var rp = getRot(part, rotDeg);

                long offW = rp.OffsetBounds.MaxX - rp.OffsetBounds.MinX;
                long offH = rp.OffsetBounds.MaxY - rp.OffsetBounds.MinY;
                if (offW <= 0 || offH <= 0 || offW > usableW || offH > usableH)
                    continue;

                var candidates = GenerateCandidates_Level1(sheet, rp, usableW, usableH, maxCandidates);
                candidates.Sort((a, b) => CandidateCompare(a, b, rp));

                for (int i = 0; i < candidates.Count; i++)
                {
                    if ((i & 127) == 0)
                        progress.ThrowIfCancelled();

                    var cand = candidates[i];

                    if (!CandidateFits(cand, rp, usableW, usableH, out var movedBBox))
                        continue;

                    var moved = TranslatePath(rp.PolyOffset, cand.InsX, cand.InsY);

                    if (OverlapsAnything(sheet, moved, movedBBox))
                        continue;

                    placement = new ContourPlacement
                    {
                        InsertX = cand.InsX,
                        InsertY = cand.InsY,
                        RotRad = rp.RotRad,
                        OffsetPolyTranslated = moved,
                        OffsetBBoxTranslated = movedBBox,
                        RotArea2Abs = rp.RotArea2Abs
                    };
                    return true;
                }
            }

            return false;
        }

        private static List<CandidateIns> GenerateCandidates_Level1(
            SheetContourState sheet,
            RotatedPoly rp,
            long usableW,
            long usableH,
            int maxCandidates)
        {
            var result = new List<CandidateIns>(Math.Min(maxCandidates, 2048));
            var seen = new HashSet<(long, long)>();

            void Add(long ix, long iy)
            {
                if (result.Count >= maxCandidates)
                    return;

                if (!seen.Add((ix, iy)))
                    return;

                long minX = ix + rp.OffsetBounds.MinX;
                long minY = iy + rp.OffsetBounds.MinY;
                long maxX = ix + rp.OffsetBounds.MaxX;
                long maxY = iy + rp.OffsetBounds.MaxY;

                if (minX < 0 || minY < 0 || maxX > usableW || maxY > usableH)
                    return;

                result.Add(new CandidateIns { InsX = ix, InsY = iy });
            }

            Add(-rp.OffsetBounds.MinX, -rp.OffsetBounds.MinY);

            var xSet = new HashSet<long> { 0 };
            var ySet = new HashSet<long> { 0 };

            foreach (var p in sheet.Placed)
            {
                xSet.Add(p.BBox.MaxX);
                ySet.Add(p.BBox.MaxY);
            }

            var xs = xSet.OrderBy(v => v).Take(140).ToList();
            var ys = ySet.OrderBy(v => v).Take(140).ToList();

            foreach (var y in ys)
            {
                foreach (var x in xs)
                {
                    Add(x - rp.OffsetBounds.MinX, y - rp.OffsetBounds.MinY);
                    if (result.Count >= maxCandidates) break;
                }
                if (result.Count >= maxCandidates) break;
            }

            if (result.Count < maxCandidates && sheet.Placed.Count > 0)
            {
                foreach (var placed in sheet.Placed)
                {
                    int n = placed.OffsetPoly.Count;
                    if (n < 3) continue;

                    int step = Math.Max(1, n / 30);

                    for (int i = 0; i < n; i += step)
                    {
                        var v = placed.OffsetPoly[i];
                        for (int a = 0; a < rp.Anchors.Length; a++)
                        {
                            var m = rp.Anchors[a];
                            Add(v.X - m.X, v.Y - m.Y);
                            if (result.Count >= maxCandidates) break;
                        }
                        if (result.Count >= maxCandidates) break;
                    }

                    if (result.Count >= maxCandidates) break;
                }
            }

            return result;
        }

        // ---------------------------
        // CONTOUR LEVEL 2 (HYBRID L2 then L1)
        // ---------------------------
        private static int NestContourLevel2_Nfp(
            List<PartDefinition> defs,
            BlockRecord modelSpace,
            double sheetWmm,
            double sheetHmm,
            double marginMm,
            double gapMm,
            double baseOriginXmm,
            double baseOriginYmm,
            string materialLabel,
            LaserCutProgressForm progress,
            int totalInstances,
            double chordMm,
            double snapMm,
            int maxCandidates,
            int maxPartners)
        {
            progress.ThrowIfCancelled();

            // hard caps
            maxCandidates = Math.Min(Math.Max(200, maxCandidates), LEVEL2_MAX_CANDIDATES_CAP);
            maxPartners = Math.Min(Math.Max(5, maxPartners), LEVEL2_MAX_PARTNERS_CAP);

            // deadline (if exceeded -> we still finish job by falling back to Level1 internally)
            long deadlineTicks = MakeDeadlineTicks(LEVEL2_JOB_TIMEOUT_SECONDS);

            int hybridTopN = ComputeHybridTopN(totalInstances);

            double boundaryBufferMm = marginMm + gapMm / 2.0;

            double usableWmm = sheetWmm - 2 * boundaryBufferMm;
            double usableHmm = sheetHmm - 2 * boundaryBufferMm;
            if (usableWmm <= 0 || usableHmm <= 0)
                throw new InvalidOperationException("Sheet is too small after margins/gap.");

            long usableW = ToInt(usableWmm);
            long usableH = ToInt(usableHmm);

            var instances = ExpandInstances(defs);
            instances.Sort((a, b) => SortByAreaDesc(a, b));

            var polyCache = new Dictionary<string, RotatedPoly>(StringComparer.OrdinalIgnoreCase);
            RotatedPoly GetRot(PartDefinition part, int rotDeg) => GetOrCreateRotated(part, rotDeg, gapMm, polyCache);

            var sheets = new List<SheetContourState>();

            SheetContourState NewSheet()
            {
                var s = new SheetContourState
                {
                    Index = sheets.Count + 1,
                    OriginXmm = baseOriginXmm + (sheets.Count) * (sheetWmm + 60.0),
                    OriginYmm = baseOriginYmm
                };

                DrawSheetOutline(s.OriginXmm, s.OriginYmm, sheetWmm, sheetHmm, modelSpace, materialLabel, s.Index, NestingMode.ContourLevel2_NFP);
                sheets.Add(s);
                return s;
            }

            var cur = NewSheet();

            int placed = 0;
            bool level2DisabledByTime = false;

            progress.ReportPlaced(0, totalInstances, sheets.Count);

            foreach (var part in instances)
            {
                progress.ThrowIfCancelled();

                bool wantLevel2 = !level2DisabledByTime && (placed < hybridTopN);

                if (wantLevel2 && Stopwatch.GetTimestamp() > deadlineTicks)
                {
                    level2DisabledByTime = true;
                    wantLevel2 = false;
                    progress.SetStatus($"Level 2 budget used ({LEVEL2_JOB_TIMEOUT_SECONDS}s) → switching to Level 1 for the rest...");
                }

                bool placedThis = false;

                for (int si = 0; si < sheets.Count; si++)
                {
                    if (wantLevel2)
                    {
                        // Try Level2 first, then Level1 fallback for this part
                        try
                        {
                            if (TryPlaceContourOnSheet_Level2Nfp(sheets[si], part, usableW, usableH, maxCandidates, maxPartners, GetRot, progress, deadlineTicks, out var placementL2))
                            {
                                CommitPlaced(modelSpace, part, sheets[si], boundaryBufferMm, placementL2);
                                placed++;
                                progress.ReportPlaced(placed, totalInstances, sheets.Count);
                                placedThis = true;
                                break;
                            }
                        }
                        catch (Level2TimeoutException)
                        {
                            level2DisabledByTime = true;
                            wantLevel2 = false;
                            progress.SetStatus($"Level 2 timeout → switching to Level 1 for the rest...");
                        }

                        if (!placedThis)
                        {
                            if (TryPlaceContourOnSheet_Level1(sheets[si], part, usableW, usableH, maxCandidates, GetRot, progress, out var placementL1))
                            {
                                CommitPlaced(modelSpace, part, sheets[si], boundaryBufferMm, placementL1);
                                placed++;
                                progress.ReportPlaced(placed, totalInstances, sheets.Count);
                                placedThis = true;
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (TryPlaceContourOnSheet_Level1(sheets[si], part, usableW, usableH, maxCandidates, GetRot, progress, out var placementL1))
                        {
                            CommitPlaced(modelSpace, part, sheets[si], boundaryBufferMm, placementL1);
                            placed++;
                            progress.ReportPlaced(placed, totalInstances, sheets.Count);
                            placedThis = true;
                            break;
                        }
                    }
                }

                if (placedThis)
                    continue;

                // Need a new sheet
                cur = NewSheet();
                progress.ReportPlaced(placed, totalInstances, sheets.Count);

                // On fresh sheet, Level1 must always work if the sheet is large enough
                // Try Level2 if still desired, else do Level1
                if (wantLevel2 && !level2DisabledByTime)
                {
                    bool ok = false;
                    try
                    {
                        ok = TryPlaceContourOnSheet_Level2Nfp(cur, part, usableW, usableH, maxCandidates, maxPartners, GetRot, progress, deadlineTicks, out var placementFreshL2);
                        if (ok)
                        {
                            CommitPlaced(modelSpace, part, cur, boundaryBufferMm, placementFreshL2);
                        }
                    }
                    catch (Level2TimeoutException)
                    {
                        level2DisabledByTime = true;
                        ok = false;
                    }

                    if (!ok)
                    {
                        if (!TryPlaceContourOnSheet_Level1(cur, part, usableW, usableH, maxCandidates, GetRot, progress, out var placementFreshL1))
                            throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

                        CommitPlaced(modelSpace, part, cur, boundaryBufferMm, placementFreshL1);
                    }
                }
                else
                {
                    if (!TryPlaceContourOnSheet_Level1(cur, part, usableW, usableH, maxCandidates, GetRot, progress, out var placementFreshL1))
                        throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

                    CommitPlaced(modelSpace, part, cur, boundaryBufferMm, placementFreshL1);
                }

                placed++;
                progress.ReportPlaced(placed, totalInstances, sheets.Count);
            }

            AddFillLabels(modelSpace, sheets, usableW, usableH, sheetWmm, sheetHmm);
            return sheets.Count;
        }

        private static void CommitPlaced(BlockRecord modelSpace, PartDefinition part, SheetContourState sheet, double boundaryBufferMm, ContourPlacement placement)
        {
            AddPlacedToDwg(modelSpace, part, sheet, boundaryBufferMm, placement.InsertX, placement.InsertY, placement.RotRad);

            // cache NFP simplified contour once
            var dec = CleanPath(DecimatePath(placement.OffsetPolyTranslated, LEVEL2_NFP_MAX_POINTS));

            sheet.Placed.Add(new PlacedContour
            {
                OffsetPoly = placement.OffsetPolyTranslated,
                OffsetPolyNfp = dec,
                BBox = placement.OffsetBBoxTranslated
            });

            sheet.PlacedCount++;
            sheet.UsedArea2Abs += placement.RotArea2Abs;
        }

        private static bool TryPlaceContourOnSheet_Level2Nfp(
            SheetContourState sheet,
            PartDefinition part,
            long usableW,
            long usableH,
            int maxCandidates,
            int maxPartners,
            Func<PartDefinition, int, RotatedPoly> getRot,
            LaserCutProgressForm progress,
            long deadlineTicks,
            out ContourPlacement placement)
        {
            placement = default;

            foreach (int rotDeg in RotationsDeg)
            {
                progress.ThrowIfCancelled();
                ThrowIfDeadlineExceeded(deadlineTicks);

                var rp = getRot(part, rotDeg);

                long offW = rp.OffsetBounds.MaxX - rp.OffsetBounds.MinX;
                long offH = rp.OffsetBounds.MaxY - rp.OffsetBounds.MinY;
                if (offW <= 0 || offH <= 0 || offW > usableW || offH > usableH)
                    continue;

                var candidates = GenerateCandidates_Level2Nfp(sheet, rp, usableW, usableH, maxCandidates, maxPartners, progress, deadlineTicks);
                candidates.Sort((a, b) => CandidateCompare(a, b, rp));

                for (int i = 0; i < candidates.Count; i++)
                {
                    if ((i & 63) == 0)
                    {
                        progress.ThrowIfCancelled();
                        ThrowIfDeadlineExceeded(deadlineTicks);
                    }

                    var cand = candidates[i];

                    if (!CandidateFits(cand, rp, usableW, usableH, out var movedBBox))
                        continue;

                    var moved = TranslatePath(rp.PolyOffset, cand.InsX, cand.InsY);

                    if (OverlapsAnything(sheet, moved, movedBBox))
                        continue;

                    placement = new ContourPlacement
                    {
                        InsertX = cand.InsX,
                        InsertY = cand.InsY,
                        RotRad = rp.RotRad,
                        OffsetPolyTranslated = moved,
                        OffsetBBoxTranslated = movedBBox,
                        RotArea2Abs = rp.RotArea2Abs
                    };
                    return true;
                }
            }

            return false;
        }

        private static List<CandidateIns> GenerateCandidates_Level2Nfp(
            SheetContourState sheet,
            RotatedPoly rp,
            long usableW,
            long usableH,
            int maxCandidates,
            int maxPartners,
            LaserCutProgressForm progress,
            long deadlineTicks)
        {
            var result = new List<CandidateIns>(Math.Min(maxCandidates, 2048));
            var seen = new HashSet<(long, long)>();

            void Add(long ix, long iy)
            {
                if (result.Count >= maxCandidates)
                    return;

                if (!seen.Add((ix, iy)))
                    return;

                long minX = ix + rp.OffsetBounds.MinX;
                long minY = iy + rp.OffsetBounds.MinY;
                long maxX = ix + rp.OffsetBounds.MaxX;
                long maxY = iy + rp.OffsetBounds.MaxY;

                if (minX < 0 || minY < 0 || maxX > usableW || maxY > usableH)
                    return;

                result.Add(new CandidateIns { InsX = ix, InsY = iy });
            }

            // bottom-left
            Add(-rp.OffsetBounds.MinX, -rp.OffsetBounds.MinY);

            if (sheet.Placed.Count == 0)
                return result;

            // small skyline fallback (bounded)
            {
                var xSet = new HashSet<long> { 0 };
                var ySet = new HashSet<long> { 0 };

                foreach (var p in sheet.Placed)
                {
                    xSet.Add(p.BBox.MaxX);
                    ySet.Add(p.BBox.MaxY);
                }

                var xs = xSet.OrderBy(v => v).Take(24).ToList();
                var ys = ySet.OrderBy(v => v).Take(24).ToList();

                int skylineBudget = Math.Min(250, maxCandidates / 2);

                foreach (var y in ys)
                {
                    foreach (var x in xs)
                    {
                        Add(x - rp.OffsetBounds.MinX, y - rp.OffsetBounds.MinY);
                        if (result.Count >= skylineBudget) break;
                    }
                    if (result.Count >= skylineBudget) break;
                }
            }

            if (result.Count >= maxCandidates)
                return result;

            // NFP candidates: MinkowskiSum(placed, -moving) using cached simplified polygons
            int partnerCount = Math.Min(maxPartners, sheet.Placed.Count);

            for (int pi = sheet.Placed.Count - 1; pi >= 0 && partnerCount > 0; pi--, partnerCount--)
            {
                if ((pi & 3) == 0)
                {
                    progress.ThrowIfCancelled();
                    ThrowIfDeadlineExceeded(deadlineTicks);
                }

                var placed = sheet.Placed[pi];
                var placedPoly = placed.OffsetPolyNfp ?? placed.OffsetPoly;
                if (placedPoly == null || placedPoly.Count < 3)
                    continue;

                Paths64 nfpPaths;
                try
                {
                    nfpPaths = MinkowskiSumSafe(placedPoly, rp.PolyOffsetNfpNeg, true);
                }
                catch
                {
                    continue;
                }

                if (nfpPaths == null || nfpPaths.Count == 0)
                    continue;

                foreach (var p in nfpPaths)
                {
                    if (p == null || p.Count < 3)
                        continue;

                    int step = Math.Max(1, p.Count / 25);
                    for (int i = 0; i < p.Count; i += step)
                    {
                        var v = p[i];
                        Add(v.X, v.Y);
                        if (result.Count >= maxCandidates) break;
                    }

                    if (result.Count >= maxCandidates) break;
                }

                if (result.Count >= maxCandidates)
                    break;
            }

            return result;
        }

        private static Path64 DecimatePath(Path64 poly, int maxPoints)
        {
            if (poly == null) return poly;
            int n = poly.Count;
            if (n <= maxPoints) return poly;

            int step = (int)Math.Ceiling(n / (double)maxPoints);
            step = Math.Max(1, step);

            var res = new Path64(Math.Min(maxPoints + 2, n));
            for (int i = 0; i < n; i += step)
                res.Add(poly[i]);

            if (res.Count < 3)
                return poly;

            return res;
        }

        // ---------------------------
        // Shared helpers
        // ---------------------------
        private static int SortByAreaDesc(PartDefinition a, PartDefinition b)
        {
            long areaA = a.OuterArea2Abs > 0 ? a.OuterArea2Abs : ToInt(a.Width) * ToInt(a.Height);
            long areaB = b.OuterArea2Abs > 0 ? b.OuterArea2Abs : ToInt(b.Width) * ToInt(b.Height);
            return areaB.CompareTo(areaA);
        }

        private static RotatedPoly GetOrCreateRotated(PartDefinition part, int rotDeg, double gapMm, Dictionary<string, RotatedPoly> cache)
        {
            string key =
                (part.BlockName ?? "") +
                "||" + rotDeg.ToString(CultureInfo.InvariantCulture) +
                "||gap:" + gapMm.ToString("0.###", CultureInfo.InvariantCulture);

            if (cache.TryGetValue(key, out var rp))
                return rp;

            Path64 basePoly = part.OuterContour0;
            if (basePoly == null || basePoly.Count < 3)
                basePoly = MakeRectPolyScaled(part.MinX, part.MinY, part.MaxX, part.MaxY);

            Path64 rotPoly = RotatePoly(basePoly, rotDeg);

            double delta = (gapMm / 2.0) * SCALE;
            Path64 offset = OffsetLargest(rotPoly, delta);
            if (offset == null || offset.Count < 3)
                offset = rotPoly;

            offset = CleanPath(offset);

            var bbox = GetBounds(offset);
            var anchors = GetAnchors(offset);
            long area2Abs = Area2Abs(rotPoly);

            // Cache NFP simplified + negated once per (part,rot,gap)
            var nfp = CleanPath(DecimatePath(offset, LEVEL2_NFP_MAX_POINTS));
            var nfpNeg = NegatePath(nfp);

            rp = new RotatedPoly
            {
                RotDeg = rotDeg,
                RotRad = rotDeg * Math.PI / 180.0,

                PolyRot = rotPoly,
                PolyOffset = offset,

                PolyOffsetNfp = nfp,
                PolyOffsetNfpNeg = nfpNeg,

                OffsetBounds = bbox,
                Anchors = anchors,
                RotArea2Abs = area2Abs
            };

            cache[key] = rp;
            return rp;
        }

        private static int CandidateCompare(CandidateIns a, CandidateIns b, RotatedPoly rp)
        {
            long aMinY = a.InsY + rp.OffsetBounds.MinY;
            long bMinY = b.InsY + rp.OffsetBounds.MinY;
            int cmp = aMinY.CompareTo(bMinY);
            if (cmp != 0) return cmp;

            long aMinX = a.InsX + rp.OffsetBounds.MinX;
            long bMinX = b.InsX + rp.OffsetBounds.MinX;
            return aMinX.CompareTo(bMinX);
        }

        private static bool CandidateFits(CandidateIns cand, RotatedPoly rp, long usableW, long usableH, out LongRect movedBBox)
        {
            long minX = cand.InsX + rp.OffsetBounds.MinX;
            long minY = cand.InsY + rp.OffsetBounds.MinY;
            long maxX = cand.InsX + rp.OffsetBounds.MaxX;
            long maxY = cand.InsY + rp.OffsetBounds.MaxY;

            movedBBox = new LongRect { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };

            if (minX < 0 || minY < 0 || maxX > usableW || maxY > usableH)
                return false;

            return true;
        }

        private static bool OverlapsAnything(SheetContourState sheet, Path64 moved, LongRect movedBBox)
        {
            foreach (var placed in sheet.Placed)
            {
                if (!RectsOverlap(movedBBox, placed.BBox))
                    continue;

                if (PolygonsOverlapAreaPositive(moved, placed.OffsetPoly))
                    return true;
            }

            return false;
        }
    }
}
