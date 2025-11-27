using System;
using System.Collections.Generic;
using System.Text;

namespace SW2026RibbonAddin
{
    /// <summary>
    /// Utilities to prepare/restore mixed Persian/Arabic (RTL) + English (LTR) text
    /// for SolidWorks' LTR-only note control.
    ///
    /// Outbound (editor -> SolidWorks):  PrepareForSolidWorks(...)
    /// Inbound  (SolidWorks -> editor):  FromSolidWorks(...)
    ///
    /// Modes:
    ///  - useRtlMarkers = true:
    ///      wrap RTL runs with classic RLE…PDF embedding marks (normally invisible in SW).
    ///  - useRtlMarkers = false (default-safe):
    ///      produce a "visual pre-compensation": keep characters within each RTL run intact,
    ///      but reverse the sequence of RTL runs across the line so final display matches
    ///      the logical order when rendered in an LTR paragraph.
    ///
    /// Optional: Arabic shaping (joins/ligatures) to avoid disconnected letters in hosts
    /// that don't shape Arabic properly.
    /// </summary>
    public static class ArabicTextUtils
    {
        private const char ZWJ = '\u200D';  // Zero Width Joiner
        private const char ZWNJ = '\u200C'; // Zero Width Non-Joiner

        // Classic BiDi embedding marks (widely supported & invisible in SW)
        private const char RLE = '\u202B';  // Right-to-Left Embedding
        private const char LRE = '\u202A';  // Left-to-Right Embedding
        private const char RLO = '\u202E';  // Right-to-Left Override
        private const char LRO = '\u202D';  // Left-to-Right Override
        private const char PDF = '\u202C';  // Pop Directional Formatting

        // Isolates & marks (we strip them inbound just in case)
        private const char LRI = '\u2066';
        private const char RLI = '\u2067';
        private const char FSI = '\u2068';
        private const char PDI = '\u2069';
        private const char LRM = '\u200E';
        private const char RLM = '\u200F';
        private const char ALM = '\u061C';

        private static bool IsBidiMarker(char c)
            => c == RLE || c == LRE || c == RLO || c == LRO || c == PDF
            || c == LRI || c == RLI || c == FSI || c == PDI
            || c == LRM || c == RLM || c == ALM;

        public static bool ContainsBidiMarkers(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (var ch in s) if (IsBidiMarker(ch)) return true;
            return false;
        }

        private static string StripBidiMarkers(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s) if (!IsBidiMarker(ch)) sb.Append(ch);
            return sb.ToString();
        }

        private enum Dir { LTR, RTL, NEUTRAL }

        private struct Run
        {
            public Dir Dir;
            public string Text;
            public Run(Dir d, string t) { Dir = d; Text = t; }
        }

        [Flags]
        private enum Join { None = 0, Prev = 1, Next = 2 }

        private struct Forms
        {
            public char Iso, Fin, Ini, Med;
            public Forms(char iso, char fin, char ini, char med)
            { Iso = iso; Fin = fin; Ini = ini; Med = med; }
        }

        // Presentation form map (subset sufficient for Persian)
        private static readonly Dictionary<char, Forms> Map = new Dictionary<char, Forms>
        {
            ['\u0621'] = new Forms('\uFE80', '\uFE80', '\uFE80', '\uFE80'),

            ['\u0622'] = new Forms('\uFE81', '\uFE82', '\uFE81', '\uFE82'),
            ['\u0623'] = new Forms('\uFE83', '\uFE84', '\uFE83', '\uFE84'),
            ['\u0625'] = new Forms('\uFE87', '\uFE88', '\uFE87', '\uFE88'),
            ['\u0627'] = new Forms('\uFE8D', '\uFE8E', '\uFE8D', '\uFE8E'),

            ['\u0628'] = new Forms('\uFE8F', '\uFE90', '\uFE91', '\uFE92'),
            ['\u062A'] = new Forms('\uFE95', '\uFE96', '\uFE97', '\uFE98'),
            ['\u062B'] = new Forms('\uFE99', '\uFE9A', '\uFE9B', '\uFE9C'),
            ['\u062C'] = new Forms('\uFE9D', '\uFE9E', '\uFE9F', '\uFEA0'),
            ['\u062D'] = new Forms('\uFEA1', '\uFEA2', '\uFEA3', '\uFEA4'),
            ['\u062E'] = new Forms('\uFEA5', '\uFEA6', '\uFEA7', '\uFEA8'),

            ['\u062F'] = new Forms('\uFEA9', '\uFEAA', '\uFEA9', '\uFEAA'),
            ['\u0630'] = new Forms('\uFEAB', '\uFEAC', '\uFEAB', '\uFEAC'),
            ['\u0631'] = new Forms('\uFEAD', '\uFEAE', '\uFEAD', '\uFEAE'),
            ['\u0632'] = new Forms('\uFEAF', '\uFEB0', '\uFEAF', '\uFEB0'),

            ['\u0633'] = new Forms('\uFEB1', '\uFEB2', '\uFEB3', '\uFEB4'),
            ['\u0634'] = new Forms('\uFEB5', '\uFEB6', '\uFEB7', '\uFEB8'),
            ['\u0635'] = new Forms('\uFEB9', '\uFEBA', '\uFEBB', '\uFEBC'),
            ['\u0636'] = new Forms('\uFEBD', '\uFEBE', '\uFEBF', '\uFEC0'),
            ['\u0637'] = new Forms('\uFEC1', '\uFEC2', '\uFEC3', '\uFEC4'),
            ['\u0638'] = new Forms('\uFEC5', '\uFEC6', '\uFEC7', '\uFEC8'),

            ['\u0639'] = new Forms('\uFEC9', '\uFECA', '\uFECB', '\uFECC'),
            ['\u063A'] = new Forms('\uFECD', '\uFECE', '\uFECF', '\uFED0'),

            ['\u0641'] = new Forms('\uFED1', '\uFED2', '\uFED3', '\uFED4'),
            ['\u0642'] = new Forms('\uFED5', '\uFED6', '\uFED7', '\uFED8'),
            ['\u0643'] = new Forms('\uFED9', '\uFEDA', '\uFEDB', '\uFEDC'),
            ['\u0644'] = new Forms('\uFEDD', '\uFEDE', '\uFEDF', '\uFEE0'),
            ['\u0645'] = new Forms('\uFEE1', '\uFEE2', '\uFEE3', '\uFEE4'),
            ['\u0646'] = new Forms('\uFEE5', '\uFEE6', '\uFEE7', '\uFEE8'),
            ['\u0647'] = new Forms('\uFEE9', '\uFEEA', '\uFEEB', '\uFEEC'),

            ['\u0648'] = new Forms('\uFEED', '\uFEEE', '\uFEED', '\uFEEE'),
            ['\u0629'] = new Forms('\uFE93', '\uFE94', '\uFE93', '\uFE94'),
            ['\u064A'] = new Forms('\uFEF1', '\uFEF2', '\uFEF3', '\uFEF4'),

            // Persian additions
            ['\u067E'] = new Forms('\uFB56', '\uFB57', '\uFB58', '\uFB59'),
            ['\u0686'] = new Forms('\uFB7A', '\uFB7B', '\uFB7C', '\uFB7D'),
            ['\u0698'] = new Forms('\uFB8A', '\uFB8B', '\uFB8A', '\uFB8B'),
            ['\u06A9'] = new Forms('\uFB8E', '\uFB8F', '\uFB90', '\uFB91'),
            ['\u06AF'] = new Forms('\uFB92', '\uFB93', '\uFB94', '\uFB95'),
            ['\u06CC'] = new Forms('\uFBFC', '\uFBFD', '\uFBFE', '\uFBFF'),
        };

        private static readonly Dictionary<char, Join> JoinType = new Dictionary<char, Join>
        {
            ['\u0621'] = Join.None,

            ['\u0622'] = Join.Prev,
            ['\u0623'] = Join.Prev,
            ['\u0625'] = Join.Prev,
            ['\u0627'] = Join.Prev,

            ['\u0628'] = Join.Prev | Join.Next,
            ['\u062A'] = Join.Prev | Join.Next,
            ['\u062B'] = Join.Prev | Join.Next,
            ['\u062C'] = Join.Prev | Join.Next,
            ['\u062D'] = Join.Prev | Join.Next,
            ['\u062E'] = Join.Prev | Join.Next,

            ['\u062F'] = Join.Prev,
            ['\u0630'] = Join.Prev,
            ['\u0631'] = Join.Prev,
            ['\u0632'] = Join.Prev,

            ['\u0633'] = Join.Prev | Join.Next,
            ['\u0634'] = Join.Prev | Join.Next,
            ['\u0635'] = Join.Prev | Join.Next,
            ['\u0636'] = Join.Prev | Join.Next,
            ['\u0637'] = Join.Prev | Join.Next,
            ['\u0638'] = Join.Prev | Join.Next,

            ['\u0639'] = Join.Prev | Join.Next,
            ['\u063A'] = Join.Prev | Join.Next,

            ['\u0641'] = Join.Prev | Join.Next,
            ['\u0642'] = Join.Prev | Join.Next,
            ['\u0643'] = Join.Prev | Join.Next,
            ['\u0644'] = Join.Prev | Join.Next,
            ['\u0645'] = Join.Prev | Join.Next,
            ['\u0646'] = Join.Prev | Join.Next,
            ['\u0647'] = Join.Prev | Join.Next,

            ['\u0648'] = Join.Prev,
            ['\u0629'] = Join.Prev,
            ['\u064A'] = Join.Prev | Join.Next,

            // Persian
            ['\u067E'] = Join.Prev | Join.Next,
            ['\u0686'] = Join.Prev | Join.Next,
            ['\u0698'] = Join.Prev, // right-joining only
            ['\u06A9'] = Join.Prev | Join.Next,
            ['\u06AF'] = Join.Prev | Join.Next,
            ['\u06CC'] = Join.Prev | Join.Next,
        };

        // Reverse map for unshaping (presentation form -> base)
        private static readonly Dictionary<char, char> PresentToBase = BuildPresentToBase();
        private static Dictionary<char, char> BuildPresentToBase()
        {
            var d = new Dictionary<char, char>();
            foreach (var kv in Map)
            {
                var baseCh = kv.Key; var f = kv.Value;
                if (f.Iso != '\u0000') d[f.Iso] = baseCh;
                if (f.Fin != '\u0000') d[f.Fin] = baseCh;
                if (f.Ini != '\u0000') d[f.Ini] = baseCh;
                if (f.Med != '\u0000') d[f.Med] = baseCh;
            }
            return d;
        }

        // Lam-Alef ligature reverse map
        private static readonly Dictionary<char, string> LigToBaseSeq = new Dictionary<char, string>
        {
            ['\uFEFB'] = "\u0644\u0627",
            ['\uFEFC'] = "\u0644\u0627",
            ['\uFEF5'] = "\u0644\u0622",
            ['\uFEF6'] = "\u0644\u0622",
            ['\uFEF7'] = "\u0644\u0623",
            ['\uFEF8'] = "\u0644\u0623",
            ['\uFEF9'] = "\u0644\u0625",
            ['\uFEFA'] = "\u0644\u0625",
        };

        private static bool IsArabicBaseLetter(char ch) => Map.ContainsKey(ch);
        private static bool IsArabicPresentationForm(char ch)
            => (ch >= '\uFB50' && ch <= '\uFDFF') || (ch >= '\uFE70' && ch <= '\uFEFF');

        private static bool IsArabicCombining(char ch)
            => (ch >= '\u064B' && ch <= '\u065F') || ch == '\u0670' || (ch >= '\u06D6' && ch <= '\u06ED');

        private static bool CanJoinPrev(char ch) => JoinType.TryGetValue(ch, out var t) && (t & Join.Prev) != 0;
        private static bool CanJoinNext(char ch) => JoinType.TryGetValue(ch, out var t) && (t & Join.Next) != 0;

        private static bool IsLatinLetter(char ch)
            => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

        private static bool IsAsciiDigit(char ch) => (ch >= '0' && ch <= '9');
        private static bool IsArabicIndicDigit(char ch) => (ch >= '\u0660' && ch <= '\u0669');   // ٠..٩
        private static bool IsEasternArabicIndicDigit(char ch) => (ch >= '\u06F0' && ch <= '\u06F9'); // ۰..۹

        // ---------------- Public API ----------------

        public static string PrepareForSolidWorks(string input, bool useRtlMarkers, bool fixDisconnected)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var lines = input.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var ln = fixDisconnected ? NormalizePunctuation(lines[i]) : lines[i];

                ln = ReorderLineForLtrHost(ln, useRtlMarkers);

                ln = fixDisconnected ? ShapeLine(ln) : ln;

                lines[i] = ln;
            }
            return string.Join("\r\n", lines);
        }

        /// <summary>
        /// Convert text coming *from* SolidWorks into logical order for the editor.
        /// - Strips BiDi markers if present.
        /// - Unshapes presentation forms to base code points.
        /// - If no markers were present (pre-compensation path), undo the visual
        ///   pre-compensation in a way that mirrors ReorderLineForLtrHost.
        /// </summary>
        public static string FromSolidWorks(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var lines = input.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var hadMarkers = ContainsBidiMarkers(lines[i]);
                var ln = StripBidiMarkers(lines[i]);
                ln = UnshapeLine(ln); // makes RTL classification robust

                if (!hadMarkers)
                {
                    // We need to undo whatever visual pre-compensation we did
                    // in ReorderLineForLtrHost when sending text to SolidWorks.
                    // For simple lines (0 or 1 LTR run), that was ReverseRtlRunSequence;
                    // for complex mixed lines (multiple LTR runs + at least one RTL run),
                    // we reversed the *run sequence* instead.

                    var runs = TokenizeWithNeutrals(ln);
                    ResolveNeutralsInPlace(runs);
                    runs = MergeAdjacentSameDir(runs);

                    int ltrCount = 0;
                    int rtlCount = 0;
                    foreach (var r in runs)
                    {
                        if (r.Dir == Dir.LTR) ltrCount++;
                        else if (r.Dir == Dir.RTL) rtlCount++;
                    }

                    if (ltrCount > 1 && rtlCount > 0)
                    {
                        // Complex mixed line: undo run-level reversal by reversing again.
                        var sb = new StringBuilder(ln.Length);
                        for (int j = runs.Count - 1; j >= 0; j--)
                            sb.Append(runs[j].Text);
                        ln = sb.ToString();
                    }
                    else
                    {
                        // Simple case: undo visual pre-compensation using the original helper.
                        ln = ReverseRtlRunSequence(ln);
                    }
                }
                // else: markers path preserves logical order; nothing more to do.

                lines[i] = ln;
            }
            return string.Join("\r\n", lines);
        }

        // ---------------- Direction & tokenization ----------------

        private static Dir ClassifyDir(char ch)
        {
            // Arabic base letters, presentation forms, combining marks & Arabic punctuation => RTL
            if (IsArabicBaseLetter(ch) || IsArabicPresentationForm(ch) || IsArabicCombining(ch)) return Dir.RTL;
            switch (ch)
            {
                case '\u061F': // Arabic ?
                case '\u060C': // Arabic ,
                case '\u061B': // Arabic ;
                    return Dir.RTL;
            }

            // Latin & digits => LTR
            if (IsLatinLetter(ch)) return Dir.LTR;
            if (IsAsciiDigit(ch) || IsArabicIndicDigit(ch) || IsEasternArabicIndicDigit(ch)) return Dir.LTR;

            return Dir.NEUTRAL; // spaces, punctuation, symbols
        }

        /// Tokenize into runs keeping neutrals separate.
        private static List<Run> TokenizeWithNeutrals(string line)
        {
            var runs = new List<Run>();
            if (string.IsNullOrEmpty(line))
            {
                runs.Add(new Run(Dir.NEUTRAL, string.Empty));
                return runs;
            }

            int i = 0;
            while (i < line.Length)
            {
                Dir d = ClassifyDir(line[i]);
                int j = i + 1;
                while (j < line.Length && ClassifyDir(line[j]) == d) j++;
                runs.Add(new Run(d, line.Substring(i, j - i)));
                i = j;
            }
            return runs;
        }

        /// Resolve neutral runs to LTR/RTL based on surrounding strong dirs (paragraph LTR).
        private static void ResolveNeutralsInPlace(List<Run> runs)
        {
            int n = runs.Count;
            for (int i = 0; i < n; i++)
            {
                if (runs[i].Dir != Dir.NEUTRAL) continue;

                // prev strong
                int p = i - 1; while (p >= 0 && runs[p].Dir == Dir.NEUTRAL) p--;
                // next strong
                int q = i + 1; while (q < n && runs[q].Dir == Dir.NEUTRAL) q++;

                if (p >= 0 && q < n)
                    runs[i] = new Run(runs[p].Dir == runs[q].Dir ? runs[p].Dir : Dir.LTR, runs[i].Text);
                else if (p >= 0)
                    runs[i] = new Run(runs[p].Dir, runs[i].Text);
                else if (q < n)
                    runs[i] = new Run(runs[q].Dir, runs[i].Text);
                else
                    runs[i] = new Run(Dir.LTR, runs[i].Text);
            }
        }

        private static List<Run> MergeAdjacentSameDir(List<Run> runs)
        {
            var merged = new List<Run>(runs.Count);
            foreach (var r in runs)
            {
                if (merged.Count > 0 && merged[merged.Count - 1].Dir == r.Dir)
                    merged[merged.Count - 1] = new Run(r.Dir, merged[merged.Count - 1].Text + r.Text);
                else
                    merged.Add(r);
            }
            return merged;
        }

        // Swap sequence of RTL runs (used by both Prepare and From paths)
        private static string ReverseRtlRunSequence(string line)
        {
            var runs = TokenizeWithNeutrals(line);
            ResolveNeutralsInPlace(runs);
            runs = MergeAdjacentSameDir(runs);

            bool hasLTR = false, hasRTL = false;
            foreach (var r in runs) { if (r.Dir == Dir.LTR) hasLTR = true; if (r.Dir == Dir.RTL) hasRTL = true; }
            if (!(hasLTR && hasRTL)) return line;

            var rtlIdx = new List<int>();
            for (int i = 0; i < runs.Count; i++) if (runs[i].Dir == Dir.RTL) rtlIdx.Add(i);
            int take = rtlIdx.Count - 1;

            var outSb = new StringBuilder(line.Length + 8);
            for (int i = 0; i < runs.Count; i++)
            {
                var r = runs[i];
                if (r.Dir == Dir.RTL)
                    outSb.Append(runs[rtlIdx[take--]].Text); // DO NOT reverse characters inside a run
                else
                    outSb.Append(r.Text);
            }
            return outSb.ToString();
        }

        // Prepare direction for SolidWorks (wrap with markers or pre-compensate)
        private static string ReorderLineForLtrHost(string line, bool useRtlMarkers)
        {
            if (useRtlMarkers)
            {
                // Wrap only RTL runs with RLE…PDF
                var runsWithMarkers = TokenizeWithNeutrals(line);
                ResolveNeutralsInPlace(runsWithMarkers);
                runsWithMarkers = MergeAdjacentSameDir(runsWithMarkers);

                var sbMarkers = new StringBuilder(line?.Length ?? 0 + 8);
                foreach (var r in runsWithMarkers)
                {
                    if (r.Dir == Dir.RTL) sbMarkers.Append(RLE).Append(r.Text).Append(PDF);
                    else sbMarkers.Append(r.Text);
                }
                return sbMarkers.ToString();
            }

            // Marker-free path.
            // We classify the line into directional runs and decide whether to
            // use the classic RTL-run swap or the run-sequence reversal used
            // for complex mixed lines (multiple LTR runs + at least one RTL run).
            var runs2 = TokenizeWithNeutrals(line);
            ResolveNeutralsInPlace(runs2);
            runs2 = MergeAdjacentSameDir(runs2);

            int ltrCount = 0;
            int rtlCount = 0;
            foreach (var r in runs2)
            {
                if (r.Dir == Dir.LTR) ltrCount++;
                else if (r.Dir == Dir.RTL) rtlCount++;
            }

            if (ltrCount > 1 && rtlCount > 0)
            {
                // Complex mixed line: pre-compensate by reversing the entire
                // run sequence (but NOT the characters inside each run).
                var sb = new StringBuilder(line.Length);
                for (int i = runs2.Count - 1; i >= 0; i--)
                    sb.Append(runs2[i].Text);
                return sb.ToString();
            }

            // Simple case (0 or 1 LTR run): use the original RTL-run swap helper.
            return ReverseRtlRunSequence(line);
        }

        // ---------------- Unshaping / Shaping / Punctuation ----------------

        private static string UnshapeLine(string line)
        {
            if (string.IsNullOrEmpty(line)) return line;
            var sb = new StringBuilder(line.Length);
            foreach (var ch in line)
            {
                if (LigToBaseSeq.TryGetValue(ch, out var seq))
                {
                    sb.Append(seq);
                }
                else if (PresentToBase.TryGetValue(ch, out var baseCh))
                {
                    sb.Append(baseCh);
                }
                else
                {
                    sb.Append(ch);
                }
            }
            return sb.ToString();
        }

        private static string NormalizePunctuation(string line)
        {
            if (string.IsNullOrEmpty(line)) return line;
            var sb = new StringBuilder(line.Length);
            foreach (var ch in line)
            {
                switch (ch)
                {
                    case '?': sb.Append('\u061F'); break; // Arabic question mark
                    case ',': sb.Append('\u060C'); break; // Arabic comma
                    case ';': sb.Append('\u061B'); break; // Arabic semicolon
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }

        private static char? LamAlefLigature(char alef, bool joinPrev)
        {
            switch (alef)
            {
                case '\u0627': return joinPrev ? '\uFEFC' : '\uFEFB';
                case '\u0622': return joinPrev ? '\uFEF6' : '\uFEF5';
                case '\u0623': return joinPrev ? '\uFEF8' : '\uFEF7';
                case '\u0625': return joinPrev ? '\uFEFA' : '\uFEF9';
                default: return null;
            }
        }

        private static int FindPrevArabicIndex(string s, int start)
        {
            for (int p = start - 1; p >= 0; p--)
            {
                char ch = s[p];
                if (ch == ZWJ) continue;
                if (ch == ZWNJ) return -1;
                if (IsArabicCombining(ch)) continue;
                if (IsArabicBaseLetter(ch)) return p;
                return -1; // barrier
            }
            return -1;
        }

        private static int FindNextArabicIndex(string s, int start)
        {
            for (int n = start + 1; n < s.Length; n++)
            {
                char ch = s[n];
                if (ch == ZWJ) continue;
                if (ch == ZWNJ) return -1;
                if (IsArabicCombining(ch)) continue;
                if (IsArabicBaseLetter(ch)) return n;
                return -1; // barrier
            }
            return -1;
        }

        private static bool CanJoinPrevHere(string s, int index)
        {
            int prevIndex = FindPrevArabicIndex(s, index);
            return prevIndex >= 0 && CanJoinNext(s[prevIndex]) && CanJoinPrev(s[index]);
        }

        private static bool CanJoinNextHere(string s, int index)
        {
            int nextIndex = FindNextArabicIndex(s, index);
            return nextIndex >= 0 && CanJoinNext(s[index]) && CanJoinPrev(s[nextIndex]);
        }

        private static string ShapeLine(string line)
        {
            if (string.IsNullOrEmpty(line)) return line;

            var sb = new StringBuilder(line.Length * 2);
            int i = 0;
            while (i < line.Length)
            {
                char cur = line[i];

                if (!IsArabicBaseLetter(cur))
                {
                    sb.Append(cur);
                    i++;
                    continue;
                }

                bool joinPrev = CanJoinPrevHere(line, i);
                bool joinNext = CanJoinNextHere(line, i);

                // Lam-Alef ligatures
                int nextIndex = FindNextArabicIndex(line, i);
                if (cur == '\u0644' && nextIndex == i + 1)
                {
                    char next = line[nextIndex];
                    if (next == '\u0627' || next == '\u0622' || next == '\u0623' || next == '\u0625')
                    {
                        var lig = LamAlefLigature(next, joinPrev);
                        if (lig.HasValue)
                        {
                            sb.Append(lig.Value);
                            i = nextIndex + 1;
                            continue;
                        }
                    }
                }

                var f = Map[cur];
                char shaped =
                    (joinPrev && joinNext) ? f.Med :
                    (joinPrev && !joinNext) ? f.Fin :
                    (!joinPrev && joinNext) ? f.Ini :
                    f.Iso;

                sb.Append(shaped);
                i++;
            }
            return sb.ToString();
        }
    }
}
