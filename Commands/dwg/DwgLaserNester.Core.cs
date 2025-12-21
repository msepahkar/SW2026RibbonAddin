using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;
using Clipper2Lib;

namespace SW2026RibbonAddin.Commands
{
	internal static partial class DwgLaserNester
	{
		// Geometry scale for Clipper (mm -> integer)
		private const long SCALE = 1000; // 0.001 mm units
		private static readonly int[] RotationsDeg = { 0, 90, 180, 270 };

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

			// bbox in mm (fallback)
			public double MinX, MinY, MaxX, MaxY;
			public double Width, Height;

			// contour (scaled)
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
			public Path64 OffsetPoly;
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

			public Path64 PolyRot;
			public Path64 PolyOffset;

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

		public static void NestFolder(string mainFolder, LaserCutRunSettings settings, bool showUi = true)
		{
			if (settings == null)
				throw new ArgumentNullException(nameof(settings));

			if (string.IsNullOrWhiteSpace(mainFolder) || !Directory.Exists(mainFolder))
				throw new DirectoryNotFoundException("Folder not found: " + mainFolder);

			var thicknessFiles = Directory.GetFiles(mainFolder, "thickness_*.dwg", SearchOption.TopDirectoryOnly)
				.Where(f =>
				{
					string n = Path.GetFileNameWithoutExtension(f) ?? "";

					// .NET Framework: no string.Contains(StringComparison)
					bool isNested = n.IndexOf("_nested", StringComparison.OrdinalIgnoreCase) >= 0;
					bool isNestLog = n.IndexOf("_nest_", StringComparison.OrdinalIgnoreCase) >= 0;

					return !isNested && !isNestLog;
				})
				.OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
				.ToList();

			if (thicknessFiles.Count == 0)
			{
				if (showUi)
				{
					MessageBox.Show(
						"No thickness_*.dwg files found in this folder.\r\nRun Combine DWG first.",
						"Laser nesting",
						MessageBoxButtons.OK,
						MessageBoxIcon.Information);
				}
				return;
			}

			var batchSummary = new StringBuilder();
			batchSummary.AppendLine("Batch nesting summary");
			batchSummary.AppendLine("Folder: " + mainFolder);
			batchSummary.AppendLine("Sheet: " + settings.DefaultSheet);
			batchSummary.AppendLine("SeparateByMaterialExact: " + settings.SeparateByMaterialExact);
			batchSummary.AppendLine("OutputOneDwgPerMaterial: " + settings.OutputOneDwgPerMaterial);
			batchSummary.AppendLine("KeepOnlyCurrentMaterialInSourcePreview: " + settings.KeepOnlyCurrentMaterialInSourcePreview);
			batchSummary.AppendLine("Mode: " + settings.Mode);
			batchSummary.AppendLine(new string('-', 70));

			foreach (var thicknessFile in thicknessFiles)
			{
				var results = NestThicknessFile(thicknessFile, settings);

				batchSummary.AppendLine(Path.GetFileName(thicknessFile));
				foreach (var r in results)
				{
					batchSummary.AppendLine($"  Material: {r.MaterialExact}");
					batchSummary.AppendLine($"  Mode: {r.Mode}");
					batchSummary.AppendLine($"  SheetsUsed: {r.SheetsUsed}, Parts: {r.TotalParts}");
					batchSummary.AppendLine($"  Output: {Path.GetFileName(r.OutputDwg)}");
				}
				batchSummary.AppendLine(new string('-', 70));
			}

			string summaryPath = Path.Combine(mainFolder, "batch_nest_summary.txt");
			File.WriteAllText(summaryPath, batchSummary.ToString(), Encoding.UTF8);

			if (showUi)
			{
				MessageBox.Show(
					"Batch nesting finished.\r\n\r\nSummary:\r\n" + summaryPath,
					"Laser nesting",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information);
			}
		}

		// Compatibility wrapper: old workflow (single file)
		public static void Nest(string sourceDwgPath, double sheetWidthMm, double sheetHeightMm)
		{
			var settings = new LaserCutRunSettings
			{
				DefaultSheet = new SheetPreset("Custom", sheetWidthMm, sheetHeightMm),
				SeparateByMaterialExact = false,
				OutputOneDwgPerMaterial = false,
				KeepOnlyCurrentMaterialInSourcePreview = false,
				Mode = NestingMode.ContourLevel1
			};

			NestThicknessFile(sourceDwgPath, settings);
		}

		public static List<NestRunResult> NestThicknessFile(string sourceDwgPath, LaserCutRunSettings settings)
		{
			if (!File.Exists(sourceDwgPath))
				throw new FileNotFoundException("DWG file not found.", sourceDwgPath);

			CadDocument firstDoc;
			using (var reader = new DwgReader(sourceDwgPath))
				firstDoc = reader.Read();

			var defsFirst = LoadPartDefinitions(firstDoc, settings).ToList();
			if (defsFirst.Count == 0)
				throw new InvalidOperationException("No plate blocks (P_*_Q#) found in: " + sourceDwgPath);

			var groups = BuildGroups(defsFirst, settings);
			var results = new List<NestRunResult>();

			foreach (var grp in groups)
			{
				string groupKey = grp.Key;
				string groupLabel = grp.Value;

				CadDocument doc;
				using (var reader = new DwgReader(sourceDwgPath))
					doc = reader.Read();

				var defs = LoadPartDefinitions(doc, settings)
					.Where(d => string.Equals(GroupKey(d.MaterialExact), groupKey, StringComparison.Ordinal))
					.ToList();

				int totalInstances = defs.Sum(d => d.Quantity);
				if (totalInstances <= 0)
					continue;

				double thicknessMm = TryGetPlateThicknessFromFileName(sourceDwgPath) ?? 0.0;

				double gapMm = 3.0;
				if (thicknessMm > gapMm) gapMm = thicknessMm;

				double marginMm = 10.0;
				if (thicknessMm > marginMm) marginMm = thicknessMm;

				var modelSpace = doc.BlockRecords["*Model_Space"];

				// Optional: keep only that material preview in output
				if (settings.SeparateByMaterialExact &&
					settings.OutputOneDwgPerMaterial &&
					settings.KeepOnlyCurrentMaterialInSourcePreview &&
					!string.Equals(groupLabel, "ALL", StringComparison.Ordinal))
				{
					var keepSet = new HashSet<string>(defs.Select(d => d.BlockName), StringComparer.OrdinalIgnoreCase);
					FilterSourcePreviewToTheseBlocks(doc, keepSet);
				}

				GetModelSpaceExtents(doc, out double srcMinX, out double srcMinY, out double srcMaxX, out double srcMaxY);

				double baseSheetOriginX = srcMinX;
				double baseSheetOriginY = srcMaxY + 200.0;

				using (var progress = new LaserCutProgressForm(totalInstances))
				{
					progress.Show();
					Application.DoEvents();

					int sheetsUsed;

					if (settings.Mode == NestingMode.FastRectangles)
					{
						sheetsUsed = NestFastRectangles(
							defs,
							modelSpace,
							settings.DefaultSheet.WidthMm,
							settings.DefaultSheet.HeightMm,
							marginMm,
							gapMm,
							baseSheetOriginX,
							baseSheetOriginY,
							groupLabel,
							progress,
							totalInstances);
					}
					else if (settings.Mode == NestingMode.ContourLevel1)
					{
						sheetsUsed = NestContourLevel1(
							defs,
							modelSpace,
							settings.DefaultSheet.WidthMm,
							settings.DefaultSheet.HeightMm,
							marginMm,
							gapMm,
							baseSheetOriginX,
							baseSheetOriginY,
							groupLabel,
							progress,
							totalInstances,
							chordMm: Math.Max(0.10, settings.ContourChordMm),
							snapMm: Math.Max(0.01, settings.ContourSnapMm),
							maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry));
					}
					else
					{
						sheetsUsed = NestContourLevel2_Nfp(
							defs,
							modelSpace,
							settings.DefaultSheet.WidthMm,
							settings.DefaultSheet.HeightMm,
							marginMm,
							gapMm,
							baseSheetOriginX,
							baseSheetOriginY,
							groupLabel,
							progress,
							totalInstances,
							chordMm: Math.Max(0.10, settings.ContourChordMm),
							snapMm: Math.Max(0.01, settings.ContourSnapMm),
							maxCandidates: Math.Max(500, settings.MaxCandidatesPerTry),
							maxPartners: Math.Max(10, settings.MaxNfpPartnersPerTry));
					}

					progress.Close();

					string dir = Path.GetDirectoryName(sourceDwgPath) ?? "";
					string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath) ?? "thickness";

					string outPath;
					if (settings.SeparateByMaterialExact && settings.OutputOneDwgPerMaterial)
					{
						string safeMat = MakeSafeFileToken(groupLabel);
						outPath = Path.Combine(dir, $"{nameNoExt}_nested_{safeMat}.dwg");
					}
					else
					{
						outPath = Path.Combine(dir, $"{nameNoExt}_nested.dwg");
					}

					using (var writer = new DwgWriter(outPath, doc))
						writer.Write();

					string logPath = Path.Combine(dir, $"{nameNoExt}_nest_log.txt");
					AppendNestLog(
						logPath,
						sourceDwgPath,
						groupLabel,
						settings.DefaultSheet.WidthMm,
						settings.DefaultSheet.HeightMm,
						thicknessMm,
						gapMm,
						marginMm,
						sheetsUsed,
						totalInstances,
						outPath,
						settings.Mode);

					results.Add(new NestRunResult
					{
						ThicknessFile = sourceDwgPath,
						MaterialExact = groupLabel,
						OutputDwg = outPath,
						SheetsUsed = sheetsUsed,
						TotalParts = totalInstances,
						Mode = settings.Mode
					});
				}
			}

			return results;
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

				// Quantity parsing: find last "_Q" and parse digits after it
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

				// Material extraction from block name tag
				string material = "UNKNOWN";
				if (TryExtractMaterialFromBlockName(name, out var mat))
					material = NormalizeMaterialLabel(mat);

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

		// ---------------------------
		// Material grouping
		// ---------------------------
		private static Dictionary<string, string> BuildGroups(List<PartDefinition> defs, LaserCutRunSettings settings)
		{
			var groups = new Dictionary<string, string>(StringComparer.Ordinal);

			if (!settings.SeparateByMaterialExact || !settings.OutputOneDwgPerMaterial)
			{
				groups["ALL"] = "ALL";
				return groups;
			}

			foreach (var d in defs)
			{
				string key = GroupKey(d.MaterialExact);
				if (!groups.ContainsKey(key))
					groups[key] = d.MaterialExact;
			}

			if (groups.Count == 0)
				groups["UNKNOWN"] = "UNKNOWN";

			return groups;
		}

		private static string GroupKey(string materialExact)
		{
			materialExact = NormalizeMaterialLabel(materialExact);
			return materialExact;
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

			// Supports patterns like:
			//   __MAT(Aluminum 6061-T6)
			//   __MAT[Aluminum 6061-T6]
			//   __MAT=Aluminum%206061-T6
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

					// stop at next "__" or "|"
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

		// ---------------------------
		// File thickness helper
		// ---------------------------
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

		// ---------------------------
		// Preview filtering (per-material output)
		// ---------------------------
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

			const double pad = 120.0;

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

		// ---------------------------
		// DWG visuals + labels
		// ---------------------------
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
				(string.IsNullOrWhiteSpace(materialLabel) || materialLabel.Equals("ALL", StringComparison.Ordinal) ? "" : $" | {materialLabel}") +
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

		// ---------------------------
		// Logging
		// ---------------------------
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
