using System;
using System.Linq;
using System.Text;

namespace SW2025RibbonAddin
{
    /// <summary>
    /// Convert text stored in a SW note (may contain bidi marks / presentation forms)
    /// back to a clean, editable Persian string.
    /// </summary>
    internal static class ArabicNoteCodec
    {
        private static readonly char[] BidiMarksAndJoiners = new[]
        {
            '\u200C', // ZWNJ
            '\u200D', // ZWJ
            '\u200E', // LRM
            '\u200F', // RLM
            '\u061C', // ALM
            '\u202A', // LRE
            '\u202B', // RLE
            '\u202D', // LRO
            '\u202E', // RLO
            '\u202C', // PDF
            '\u2066', // LRI
            '\u2067', // RLI
            '\u2068', // FSI
            '\u2069'  // PDI
        };

        public static string DecodeFromNote(string textInNote)
        {
            if (string.IsNullOrEmpty(textInNote)) return string.Empty;

            // Strip bidi/zero-width markers
            var noMarks = new string(textInNote.Where(c => Array.IndexOf(BidiMarksAndJoiners, c) < 0).ToArray());

            // Normalize Arabic Presentation Forms back to base letters
            string normalized = noMarks.Normalize(NormalizationForm.FormKC);

            return normalized.Trim();
        }
    }
}
