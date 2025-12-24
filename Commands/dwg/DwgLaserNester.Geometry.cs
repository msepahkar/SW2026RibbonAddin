using ACadSharp.Entities;
using ACadSharp.Tables;
using Clipper2Lib;
using CSMath;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
// Avoid ambiguity with ACadSharp.Entities.ClipType
using ClipperClipType = Clipper2Lib.ClipType;

namespace SW2026RibbonAddin.Commands
{
	internal static partial class DwgLaserNester
	{
		// cached MinkowskiSum via reflection (Clipper2 version safe)
		private static MethodInfo _miMinkowskiSum;

		private static long ToInt(double mm) => (long)Math.Round(mm * SCALE);

		private static Point64 ToP64(XYZ p) => new Point64(ToInt(p.X), ToInt(p.Y));
		private static Point64 ToP64(double x, double y) => new Point64(ToInt(x), ToInt(y));

		private static Point64 Snap(Point64 p, double snapMm)
		{
			long grid = Math.Max(1, (long)Math.Round(snapMm * SCALE));
			long sx = (long)Math.Round((double)p.X / grid) * grid;
			long sy = (long)Math.Round((double)p.Y / grid) * grid;
			return new Point64(sx, sy);
		}

		private static Path64 CleanPath(Path64 path)
		{
			if (path == null || path.Count == 0)
				return path;

			var res = new Path64();
			Point64 prev = path[0];
			res.Add(prev);

			for (int i = 1; i < path.Count; i++)
			{
				var cur = path[i];
				if (cur.X == prev.X && cur.Y == prev.Y)
					continue;

				res.Add(cur);
				prev = cur;
			}

			if (res.Count > 1 && res[0].X == res[res.Count - 1].X && res[0].Y == res[res.Count - 1].Y)
				res.RemoveAt(res.Count - 1);

			return res;
		}

		private static Path64 MakeRectPolyScaled(double minX, double minY, double maxX, double maxY)
		{
			long x1 = ToInt(minX);
			long y1 = ToInt(minY);
			long x2 = ToInt(maxX);
			long y2 = ToInt(maxY);

			return new Path64
			{
				new Point64(x1, y1),
				new Point64(x2, y1),
				new Point64(x2, y2),
				new Point64(x1, y2)
			};
		}

		private static Path64 RotatePoly(Path64 p, int rotDeg)
		{
			if (p == null) return null;
			rotDeg = ((rotDeg % 360) + 360) % 360;

			var r = new Path64(p.Count);

			foreach (var pt in p)
			{
				long x = pt.X;
				long y = pt.Y;

				switch (rotDeg)
				{
					case 0: r.Add(new Point64(x, y)); break;
					case 90: r.Add(new Point64(-y, x)); break;
					case 180: r.Add(new Point64(-x, -y)); break;
					case 270: r.Add(new Point64(y, -x)); break;
					default:
						double rad = rotDeg * Math.PI / 180.0;
						long xr = (long)Math.Round(x * Math.Cos(rad) - y * Math.Sin(rad));
						long yr = (long)Math.Round(x * Math.Sin(rad) + y * Math.Cos(rad));
						r.Add(new Point64(xr, yr));
						break;
				}
			}

			return r;
		}

		// Rotate a polygon by an arbitrary angle (radians) around the origin.
		// NOTE: points are in SCALE units (0.001mm). We round back to long.
		private static Path64 RotatePolyRad(Path64 p, double rad)
		{
			if (p == null)
				return null;

			// Fast-path for 0, 90, 180, 270-ish angles to keep results stable.
			// We only use this when very close to those angles.
			double deg = rad * 180.0 / Math.PI;
			double degNorm = ((deg % 360.0) + 360.0) % 360.0;
			if (Math.Abs(degNorm - 0.0) < 1e-9) return RotatePoly(p, 0);
			if (Math.Abs(degNorm - 90.0) < 1e-9) return RotatePoly(p, 90);
			if (Math.Abs(degNorm - 180.0) < 1e-9) return RotatePoly(p, 180);
			if (Math.Abs(degNorm - 270.0) < 1e-9) return RotatePoly(p, 270);

			double c = Math.Cos(rad);
			double s = Math.Sin(rad);

			var r = new Path64(p.Count);
			foreach (var pt in p)
			{
				double x = pt.X;
				double y = pt.Y;
				long xr = (long)Math.Round(x * c - y * s);
				long yr = (long)Math.Round(x * s + y * c);
				r.Add(new Point64(xr, yr));
			}

			return r;
		}

		private static Path64 SnapPath(Path64 p, double snapMm)
		{
			if (p == null || p.Count == 0)
				return p;

			long grid = Math.Max(1, (long)Math.Round(snapMm * SCALE));
			if (grid <= 1)
				return p;

			var r = new Path64(p.Count);
			foreach (var pt in p)
			{
				long sx = (long)Math.Round((double)pt.X / grid) * grid;
				long sy = (long)Math.Round((double)pt.Y / grid) * grid;
				r.Add(new Point64(sx, sy));
			}

			return CleanPath(r);
		}

		private static Path64 OffsetLargest(Path64 poly, double deltaScaled)
		{
			if (poly == null || poly.Count < 3)
				return null;

			var co = new ClipperOffset();
			co.AddPath(poly, JoinType.Round, EndType.Polygon);

			var sol = new Paths64();
			co.Execute(deltaScaled, sol);

			if (sol == null || sol.Count == 0)
				return null;

			Path64 best = null;
			long bestArea = 0;

			foreach (var p in sol)
			{
				long a2 = Area2Abs(p);
				if (a2 > bestArea)
				{
					bestArea = a2;
					best = p;
				}
			}

			return best;
		}

		private static LongRect GetBounds(Path64 p)
		{
			long minX = long.MaxValue, minY = long.MaxValue;
			long maxX = long.MinValue, maxY = long.MinValue;

			foreach (var pt in p)
			{
				if (pt.X < minX) minX = pt.X;
				if (pt.Y < minY) minY = pt.Y;
				if (pt.X > maxX) maxX = pt.X;
				if (pt.Y > maxY) maxY = pt.Y;
			}

			return new LongRect { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };
		}

		private static Point64[] GetAnchors(Path64 p)
		{
			Point64 bl = p[0], br = p[0], tl = p[0], tr = p[0];

			foreach (var pt in p)
			{
				if (pt.Y < bl.Y || (pt.Y == bl.Y && pt.X < bl.X)) bl = pt;
				if (pt.Y < br.Y || (pt.Y == br.Y && pt.X > br.X)) br = pt;
				if (pt.Y > tl.Y || (pt.Y == tl.Y && pt.X < tl.X)) tl = pt;
				if (pt.Y > tr.Y || (pt.Y == tr.Y && pt.X > tr.X)) tr = pt;
			}

			return new[] { bl, br, tl, tr };
		}

		private static Path64 TranslatePath(Path64 p, long dx, long dy)
		{
			var r = new Path64(p.Count);
			foreach (var pt in p)
				r.Add(new Point64(pt.X + dx, pt.Y + dy));
			return r;
		}

		private static bool RectsOverlap(LongRect a, LongRect b)
		{
			return !(a.MaxX <= b.MinX || b.MaxX <= a.MinX || a.MaxY <= b.MinY || b.MaxY <= a.MinY);
		}

		private static bool PolygonsOverlapAreaPositive(Path64 a, Path64 b)
		{
			var clipper = new Clipper64();
			clipper.AddSubject(a);
			clipper.AddClip(b);

			var sol = new Paths64();
			clipper.Execute(ClipperClipType.Intersection, FillRule.NonZero, sol);

			if (sol == null || sol.Count == 0)
				return false;

			foreach (var p in sol)
			{
				if (Area2Abs(p) > 0)
					return true;
			}

			return false;
		}

		private static long Area2Abs(Path64 p)
		{
			if (p == null || p.Count < 3)
				return 0;

			long sum = 0;
			int n = p.Count;

			for (int i = 0; i < n; i++)
			{
				var a = p[i];
				var b = p[(i + 1) % n];
				sum += a.X * b.Y - b.X * a.Y;
			}

			return Math.Abs(sum);
		}

		// ==============================
		// Contour extraction from Block
		// ==============================
		private static Path64 ExtractOuterContourScaled(BlockRecord block, double chordMm, double snapMm)
		{
			if (block == null)
				return null;

			chordMm = Math.Max(0.10, chordMm);
			snapMm = Math.Max(0.01, snapMm);

			var segs = new List<(Point64 A, Point64 B)>();

			foreach (var ent in block.Entities)
			{
				if (ent == null) continue;

				if (ent is Line ln)
				{
					segs.Add((Snap(ToP64(ln.StartPoint), snapMm), Snap(ToP64(ln.EndPoint), snapMm)));
				}
				else if (ent is Arc arc)
				{
					AddArcSegments(segs, arc.Center, arc.Radius, arc.StartAngle, arc.EndAngle, chordMm, snapMm);
				}
				else if (ent is Circle cir)
				{
					AddCircleSegments(segs, cir.Center, cir.Radius, chordMm, snapMm);
				}
				else
				{
					if (TryAddPolylineSegments(ent, segs, chordMm, snapMm))
						continue;
				}
			}

			if (segs.Count < 3)
				return null;

			var loops = BuildClosedLoops(segs);
			if (loops.Count > 0)
			{
				Path64 best = null;
				long bestArea = 0;
				foreach (var loop in loops)
				{
					long a2 = Area2Abs(loop);
					if (a2 > bestArea)
					{
						bestArea = a2;
						best = loop;
					}
				}
				return best;
			}

			// fallback: convex hull
			var pts = new List<Point64>(segs.Count * 2);
			foreach (var s in segs)
			{
				pts.Add(s.A);
				pts.Add(s.B);
			}

			return ConvexHull(pts);
		}

		private static List<Path64> BuildClosedLoops(List<(Point64 A, Point64 B)> segs)
		{
			var loops = new List<Path64>();
			if (segs == null || segs.Count == 0)
				return loops;

			var adj = new Dictionary<(long, long), List<int>>();
			var used = new bool[segs.Count];

			(long, long) Key(Point64 p) => (p.X, p.Y);

			for (int i = 0; i < segs.Count; i++)
			{
				var s = segs[i];
				var kA = Key(s.A);
				var kB = Key(s.B);

				if (!adj.TryGetValue(kA, out var la)) { la = new List<int>(); adj[kA] = la; }
				la.Add(i);

				if (!adj.TryGetValue(kB, out var lb)) { lb = new List<int>(); adj[kB] = lb; }
				lb.Add(i);
			}

			for (int i = 0; i < segs.Count; i++)
			{
				if (used[i]) continue;

				var s0 = segs[i];
				var start = s0.A;
				var startK = Key(start);

				var path = new Path64();
				path.Add(start);

				Point64 cur = s0.B;
				var curK = Key(cur);

				used[i] = true;
				path.Add(cur);

				var prevK = startK;

				int guard = 0;
				while (curK != startK && guard++ < segs.Count + 10)
				{
					if (!adj.TryGetValue(curK, out var incident))
						break;

					int nextSeg = -1;

					foreach (int si in incident)
					{
						if (used[si]) continue;

						var s = segs[si];
						var aK = Key(s.A);
						var bK = Key(s.B);

						var otherK = (aK == curK) ? bK : (bK == curK ? aK : curK);

						if (otherK != prevK)
						{
							nextSeg = si;
							break;
						}

						if (nextSeg < 0)
							nextSeg = si;
					}

					if (nextSeg < 0)
						break;

					used[nextSeg] = true;

					var ns = segs[nextSeg];
					var aK2 = Key(ns.A);
					var bK2 = Key(ns.B);

					Point64 nextPt;
					(long, long) nextK;

					if (aK2 == curK)
					{
						nextPt = ns.B;
						nextK = bK2;
					}
					else
					{
						nextPt = ns.A;
						nextK = aK2;
					}

					if (path.Count == 0 || path[path.Count - 1].X != nextPt.X || path[path.Count - 1].Y != nextPt.Y)
						path.Add(nextPt);

					prevK = curK;
					curK = nextK;
					cur = nextPt;
				}

				if (curK == startK && path.Count >= 4)
				{
					if (path.Count > 1 && path[path.Count - 1].X == path[0].X && path[path.Count - 1].Y == path[0].Y)
						path.RemoveAt(path.Count - 1);

					path = CleanPath(path);

					if (path != null && path.Count >= 3)
						loops.Add(path);
				}
			}

			return loops;
		}

		private static bool TryAddPolylineSegments(Entity ent, List<(Point64 A, Point64 B)> segs, double chordMm, double snapMm)
		{
			try
			{
				var t = ent.GetType();
				string tn = t.Name ?? "";

				// .NET Framework: no string.Contains(StringComparison)
				if (tn.IndexOf("Polyline", StringComparison.OrdinalIgnoreCase) < 0)
					return false;

				var vertsProp = t.GetProperty("Vertices");
				var vertsObj = vertsProp?.GetValue(ent);
				if (vertsObj == null)
					return false;

				var vertsEnum = vertsObj as System.Collections.IEnumerable;
				if (vertsEnum == null)
					return false;

				var verts = new List<(double X, double Y, double Bulge)>();
				foreach (var v in vertsEnum)
				{
					if (TryGetVertexXYB(v, out double x, out double y, out double b))
						verts.Add((x, y, b));
				}

				if (verts.Count < 2)
					return false;

				bool closed = false;
				var closedProp = t.GetProperty("IsClosed") ?? t.GetProperty("Closed");
				if (closedProp != null && closedProp.PropertyType == typeof(bool))
					closed = (bool)closedProp.GetValue(ent);

				int count = verts.Count;
				int last = closed ? count : count - 1;

				for (int i = 0; i < last; i++)
				{
					var v1 = verts[i];
					var v2 = verts[(i + 1) % count];

					if (Math.Abs(v1.Bulge) < 1e-12)
					{
						segs.Add((Snap(ToP64(v1.X, v1.Y), snapMm), Snap(ToP64(v2.X, v2.Y), snapMm)));
					}
					else
					{
						AddBulgeArcSegments(segs, v1, v2, chordMm, snapMm);
					}
				}

				return true;
			}
			catch
			{
				return false;
			}
		}

		private static bool TryGetVertexXYB(object v, out double x, out double y, out double bulge)
		{
			x = y = 0.0;
			bulge = 0.0;

			if (v == null) return false;

			try
			{
				var t = v.GetType();

				var pb = t.GetProperty("Bulge");
				if (pb != null)
				{
					object bv = pb.GetValue(v);
					if (bv is double bd) bulge = bd;
				}

				var px = t.GetProperty("X");
				var py = t.GetProperty("Y");

				if (px != null && py != null)
				{
					x = Convert.ToDouble(px.GetValue(v), CultureInfo.InvariantCulture);
					y = Convert.ToDouble(py.GetValue(v), CultureInfo.InvariantCulture);
					return true;
				}

				var ploc = t.GetProperty("Location") ?? t.GetProperty("Point");
				if (ploc != null)
				{
					var loc = ploc.GetValue(v);
					if (loc != null)
					{
						var lt = loc.GetType();
						var lx = lt.GetProperty("X");
						var ly = lt.GetProperty("Y");
						if (lx != null && ly != null)
						{
							x = Convert.ToDouble(lx.GetValue(loc), CultureInfo.InvariantCulture);
							y = Convert.ToDouble(ly.GetValue(loc), CultureInfo.InvariantCulture);
							return true;
						}
					}
				}
			}
			catch { }

			return false;
		}

		private static void AddBulgeArcSegments(List<(Point64 A, Point64 B)> segs, (double X, double Y, double Bulge) v1, (double X, double Y, double Bulge) v2, double chordMm, double snapMm)
		{
			double b = v1.Bulge;
			if (Math.Abs(b) < 1e-12)
			{
				segs.Add((Snap(ToP64(v1.X, v1.Y), snapMm), Snap(ToP64(v2.X, v2.Y), snapMm)));
				return;
			}

			double x1 = v1.X, y1 = v1.Y;
			double x2 = v2.X, y2 = v2.Y;

			double dx = x2 - x1;
			double dy = y2 - y1;
			double L = Math.Sqrt(dx * dx + dy * dy);
			if (L < 1e-9)
				return;

			double theta = 4.0 * Math.Atan(b); // signed
			double sinHalf = Math.Sin(theta / 2.0);
			if (Math.Abs(sinHalf) < 1e-12)
			{
				segs.Add((Snap(ToP64(x1, y1), snapMm), Snap(ToP64(x2, y2), snapMm)));
				return;
			}

			double R = L / (2.0 * sinHalf);
			double Rabs = Math.Abs(R);

			double mx = (x1 + x2) / 2.0;
			double my = (y1 + y2) / 2.0;

			double d = Math.Sqrt(Math.Max(0.0, Rabs * Rabs - (L / 2.0) * (L / 2.0)));

			double nx = -dy / L;
			double ny = dx / L;

			double sign = b >= 0 ? 1.0 : -1.0;

			double cx = mx + sign * nx * d;
			double cy = my + sign * ny * d;

			double a1 = Math.Atan2(y1 - cy, x1 - cx);

			int segCount = Math.Max(8, (int)Math.Ceiling((Rabs * Math.Abs(theta)) / Math.Max(0.10, chordMm)));
			segCount = Math.Min(segCount, 720);

			double step = theta / segCount;

			Point64 prev = Snap(ToP64(x1, y1), snapMm);
			for (int i = 1; i <= segCount; i++)
			{
				double ang = a1 + step * i;
				double px = cx + Rabs * Math.Cos(ang);
				double py = cy + Rabs * Math.Sin(ang);

				var cur = Snap(ToP64(px, py), snapMm);
				segs.Add((prev, cur));
				prev = cur;
			}
		}

		private static void AddArcSegments(List<(Point64 A, Point64 B)> segs, XYZ center, double radius, double startAngle, double endAngle, double chordMm, double snapMm)
		{
			double sa = DegreesToRadiansIfNeeded(startAngle);
			double ea = DegreesToRadiansIfNeeded(endAngle);

			double sweep = ea - sa;
			while (sweep < 0) sweep += 2.0 * Math.PI;
			if (sweep <= 1e-12) sweep = 2.0 * Math.PI;

			double r = Math.Abs(radius);
			if (r <= 1e-9) return;

			int segCount = Math.Max(8, (int)Math.Ceiling((r * sweep) / Math.Max(0.10, chordMm)));
			segCount = Math.Min(segCount, 1440);

			Point64 prev = Snap(ToP64(center.X + r * Math.Cos(sa), center.Y + r * Math.Sin(sa)), snapMm);

			for (int i = 1; i <= segCount; i++)
			{
				double ang = sa + sweep * i / segCount;
				var cur = Snap(ToP64(center.X + r * Math.Cos(ang), center.Y + r * Math.Sin(ang)), snapMm);
				segs.Add((prev, cur));
				prev = cur;
			}
		}

		private static void AddCircleSegments(List<(Point64 A, Point64 B)> segs, XYZ center, double radius, double chordMm, double snapMm)
		{
			double r = Math.Abs(radius);
			if (r <= 1e-9) return;

			double sweep = 2.0 * Math.PI;
			int segCount = Math.Max(16, (int)Math.Ceiling((r * sweep) / Math.Max(0.10, chordMm)));
			segCount = Math.Min(segCount, 2880);

			Point64 first = Snap(ToP64(center.X + r, center.Y), snapMm);
			Point64 prev = first;

			for (int i = 1; i <= segCount; i++)
			{
				double ang = sweep * i / segCount;
				var cur = Snap(ToP64(center.X + r * Math.Cos(ang), center.Y + r * Math.Sin(ang)), snapMm);
				segs.Add((prev, cur));
				prev = cur;
			}
		}

		private static double DegreesToRadiansIfNeeded(double angle)
		{
			if (Math.Abs(angle) > 10.0)
				return angle * Math.PI / 180.0;
			return angle;
		}

		private static Path64 ConvexHull(List<Point64> pts)
		{
			if (pts == null)
				return null;

			var uniq = pts.Distinct().ToList();
			if (uniq.Count < 3)
				return null;

			uniq.Sort((a, b) =>
			{
				int c = a.X.CompareTo(b.X);
				if (c != 0) return c;
				return a.Y.CompareTo(b.Y);
			});

			long Cross(Point64 o, Point64 a, Point64 b)
				=> (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X);

			var lower = new List<Point64>();
			foreach (var p in uniq)
			{
				while (lower.Count >= 2 && Cross(lower[lower.Count - 2], lower[lower.Count - 1], p) <= 0)
					lower.RemoveAt(lower.Count - 1);
				lower.Add(p);
			}

			var upper = new List<Point64>();
			for (int i = uniq.Count - 1; i >= 0; i--)
			{
				var p = uniq[i];
				while (upper.Count >= 2 && Cross(upper[upper.Count - 2], upper[upper.Count - 1], p) <= 0)
					upper.RemoveAt(upper.Count - 1);
				upper.Add(p);
			}

			lower.RemoveAt(lower.Count - 1);
			upper.RemoveAt(upper.Count - 1);

			var hull = new Path64();
			hull.AddRange(lower);
			hull.AddRange(upper);

			return hull.Count >= 3 ? hull : null;
		}

		// ==============================
		// Minkowski helpers (Level 2)
		// ==============================
		private static Path64 NegatePath(Path64 p)
		{
			if (p == null) return null;

			var r = new Path64(p.Count);
			foreach (var pt in p)
				r.Add(new Point64(-pt.X, -pt.Y));

			r.Reverse();
			return r;
		}

		private static Paths64 MinkowskiSumSafe(Path64 a, Path64 b, bool closed)
		{
			if (_miMinkowskiSum == null)
			{
				var asm = typeof(Clipper64).Assembly;

				foreach (var t in asm.GetTypes())
				{
					foreach (var m in t.GetMethods(BindingFlags.Public | BindingFlags.Static))
					{
						if (!string.Equals(m.Name, "MinkowskiSum", StringComparison.Ordinal))
							continue;

						var ps = m.GetParameters();
						if (ps.Length != 3) continue;
						if (ps[0].ParameterType != typeof(Path64)) continue;
						if (ps[1].ParameterType != typeof(Path64)) continue;
						if (ps[2].ParameterType != typeof(bool)) continue;
						if (m.ReturnType != typeof(Paths64)) continue;

						_miMinkowskiSum = m;
						break;
					}
					if (_miMinkowskiSum != null) break;
				}

				if (_miMinkowskiSum == null)
					throw new InvalidOperationException("Clipper2 MinkowskiSum(Path64, Path64, bool) not found. Check Clipper2Lib version.");
			}

			return (Paths64)_miMinkowskiSum.Invoke(null, new object[] { a, b, closed });
		}
	}
}
