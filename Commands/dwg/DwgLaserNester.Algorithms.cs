using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using ACadSharp.Entities;
using ACadSharp.Tables;
using Clipper2Lib;

namespace SW2026RibbonAddin.Commands
{
	internal static partial class DwgLaserNester
	{
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

			foreach (var part in instances)
			{
				while (true)
				{
					if (TryPlaceRect(cur, part, gapMm, modelSpace))
					{
						placed++;
						progress.Step($"Placed {placed}/{totalInstances} (Fast rectangles)");
						break;
					}

					cur = NewSheet();
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
				InsertPoint = new CSMath.XYZ(insertX, insertY, 0.0),
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

			foreach (var part in instances)
			{
				bool placedThis = false;

				for (int si = 0; si < sheets.Count; si++)
				{
					if (TryPlaceContourOnSheet_Level1(sheets[si], part, usableW, usableH, maxCandidates, GetRot, out var placement))
					{
						AddPlacedToDwg(modelSpace, part, sheets[si], boundaryBufferMm, placement.InsertX, placement.InsertY, placement.RotRad);
						sheets[si].Placed.Add(new PlacedContour { OffsetPoly = placement.OffsetPolyTranslated, BBox = placement.OffsetBBoxTranslated });
						sheets[si].PlacedCount++;
						sheets[si].UsedArea2Abs += placement.RotArea2Abs;

						placed++;
						progress.Step($"Placed {placed}/{totalInstances} (Contour L1)");
						placedThis = true;
						break;
					}
				}

				if (placedThis)
					continue;

				cur = NewSheet();

				if (!TryPlaceContourOnSheet_Level1(cur, part, usableW, usableH, maxCandidates, GetRot, out var placement2))
					throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

				AddPlacedToDwg(modelSpace, part, cur, boundaryBufferMm, placement2.InsertX, placement2.InsertY, placement2.RotRad);
				cur.Placed.Add(new PlacedContour { OffsetPoly = placement2.OffsetPolyTranslated, BBox = placement2.OffsetBBoxTranslated });
				cur.PlacedCount++;
				cur.UsedArea2Abs += placement2.RotArea2Abs;

				placed++;
				progress.Step($"Placed {placed}/{totalInstances} (Contour L1)");
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
			out ContourPlacement placement)
		{
			placement = default;

			foreach (int rotDeg in RotationsDeg)
			{
				var rp = getRot(part, rotDeg);

				long offW = rp.OffsetBounds.MaxX - rp.OffsetBounds.MinX;
				long offH = rp.OffsetBounds.MaxY - rp.OffsetBounds.MinY;
				if (offW <= 0 || offH <= 0 || offW > usableW || offH > usableH)
					continue;

				var candidates = GenerateCandidates_Level1(sheet, rp, usableW, usableH, maxCandidates);
				candidates.Sort((a, b) => CandidateCompare(a, b, rp));

				foreach (var cand in candidates)
				{
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

		private static List<CandidateIns> GenerateCandidates_Level1(SheetContourState sheet, RotatedPoly rp, long usableW, long usableH, int maxCandidates)
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
		// CONTOUR LEVEL 2 (NFP)
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

			foreach (var part in instances)
			{
				bool placedThis = false;

				for (int si = 0; si < sheets.Count; si++)
				{
					if (TryPlaceContourOnSheet_Level2Nfp(sheets[si], part, usableW, usableH, maxCandidates, maxPartners, GetRot, out var placement))
					{
						AddPlacedToDwg(modelSpace, part, sheets[si], boundaryBufferMm, placement.InsertX, placement.InsertY, placement.RotRad);
						sheets[si].Placed.Add(new PlacedContour { OffsetPoly = placement.OffsetPolyTranslated, BBox = placement.OffsetBBoxTranslated });
						sheets[si].PlacedCount++;
						sheets[si].UsedArea2Abs += placement.RotArea2Abs;

						placed++;
						progress.Step($"Placed {placed}/{totalInstances} (Contour L2 NFP)");
						placedThis = true;
						break;
					}
				}

				if (placedThis)
					continue;

				cur = NewSheet();

				if (!TryPlaceContourOnSheet_Level2Nfp(cur, part, usableW, usableH, maxCandidates, maxPartners, GetRot, out var placement2))
					throw new InvalidOperationException("Failed to place a part even on a fresh sheet. Sheet too small?");

				AddPlacedToDwg(modelSpace, part, cur, boundaryBufferMm, placement2.InsertX, placement2.InsertY, placement2.RotRad);
				cur.Placed.Add(new PlacedContour { OffsetPoly = placement2.OffsetPolyTranslated, BBox = placement2.OffsetBBoxTranslated });
				cur.PlacedCount++;
				cur.UsedArea2Abs += placement2.RotArea2Abs;

				placed++;
				progress.Step($"Placed {placed}/{totalInstances} (Contour L2 NFP)");
			}

			AddFillLabels(modelSpace, sheets, usableW, usableH, sheetWmm, sheetHmm);
			return sheets.Count;
		}

		private static bool TryPlaceContourOnSheet_Level2Nfp(
			SheetContourState sheet,
			PartDefinition part,
			long usableW,
			long usableH,
			int maxCandidates,
			int maxPartners,
			Func<PartDefinition, int, RotatedPoly> getRot,
			out ContourPlacement placement)
		{
			placement = default;

			foreach (int rotDeg in RotationsDeg)
			{
				var rp = getRot(part, rotDeg);

				long offW = rp.OffsetBounds.MaxX - rp.OffsetBounds.MinX;
				long offH = rp.OffsetBounds.MaxY - rp.OffsetBounds.MinY;
				if (offW <= 0 || offH <= 0 || offW > usableW || offH > usableH)
					continue;

				var candidates = GenerateCandidates_Level2Nfp(sheet, rp, usableW, usableH, maxCandidates, maxPartners);
				candidates.Sort((a, b) => CandidateCompare(a, b, rp));

				foreach (var cand in candidates)
				{
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
			int maxPartners)
		{
			var result = new List<CandidateIns>(Math.Min(maxCandidates, 4096));
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

			if (sheet.Placed.Count == 0)
				return result;

			// small grid fallback
			{
				var xSet = new HashSet<long> { 0 };
				var ySet = new HashSet<long> { 0 };

				foreach (var p in sheet.Placed)
				{
					xSet.Add(p.BBox.MaxX);
					ySet.Add(p.BBox.MaxY);
				}

				var xs = xSet.OrderBy(v => v).Take(100).ToList();
				var ys = ySet.OrderBy(v => v).Take(100).ToList();

				foreach (var y in ys)
				{
					foreach (var x in xs)
					{
						Add(x - rp.OffsetBounds.MinX, y - rp.OffsetBounds.MinY);
						if (result.Count >= maxCandidates) break;
					}
					if (result.Count >= maxCandidates) break;
				}
			}

			int partnerCount = Math.Min(maxPartners, sheet.Placed.Count);

			for (int pi = sheet.Placed.Count - 1; pi >= 0 && partnerCount > 0; pi--, partnerCount--)
			{
				var placed = sheet.Placed[pi];
				if (placed.OffsetPoly == null || placed.OffsetPoly.Count < 3)
					continue;

				var negA = NegatePath(rp.PolyOffset);

				Paths64 nfpPaths;
				try
				{
					nfpPaths = MinkowskiSumSafe(placed.OffsetPoly, negA, true);
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

					int step = Math.Max(1, p.Count / 35);
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

		// ---------------------------
		// Shared placement helpers
		// ---------------------------
		private static int SortByAreaDesc(PartDefinition a, PartDefinition b)
		{
			long areaA = a.OuterArea2Abs > 0 ? a.OuterArea2Abs : ToInt(a.Width) * ToInt(a.Height);
			long areaB = b.OuterArea2Abs > 0 ? b.OuterArea2Abs : ToInt(b.Width) * ToInt(b.Height);
			return areaB.CompareTo(areaA);
		}

		private static RotatedPoly GetOrCreateRotated(PartDefinition part, int rotDeg, double gapMm, Dictionary<string, RotatedPoly> cache)
		{
			string key = part.BlockName + "||" + rotDeg.ToString(CultureInfo.InvariantCulture) + "||gap:" + gapMm.ToString("0.###", CultureInfo.InvariantCulture);

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

			rp = new RotatedPoly
			{
				RotDeg = rotDeg,
				RotRad = rotDeg * Math.PI / 180.0,
				PolyRot = rotPoly,
				PolyOffset = offset,
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
