using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin
{
    /// <summary>
    /// Simple container for all fastener metadata we care about.
    /// </summary>
    internal sealed class FastenerInitialValues
    {
        public string Family { get; set; }              // "Bolt", "Washer", "Nut"
        public string PartNumber { get; set; }
        public string Standard { get; set; }
        public string NominalDiameter { get; set; }     // stored as text, e.g. "4"
        public string Length { get; set; }              // bolts only, text "10"
        public string StrengthClass { get; set; }       // "8.8", "10.9", etc.

        public string BoltType { get; set; }            // HexBolt, SocketHeadCap, ...
        public string WasherType { get; set; }          // PlainWasher, SpringWasher
        public string NutType { get; set; }             // HexNut, LockNut, ...

        public string OuterDiameter { get; set; }       // washers
        public string Thickness { get; set; }           // washers
        public string Height { get; set; }              // optional, nuts

        public string Description { get; set; }
        public string Material { get; set; }
    }

    /// <summary>
    /// Helper to read/write fastener custom properties and infer defaults
    /// from filename / Toolbox naming.
    /// </summary>
    internal static class FastenerPropertyHelper
    {
        public static FastenerInitialValues BuildInitialValues(IModelDoc2 model)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));

            var values = new FastenerInitialValues
            {
                PartNumber = SafeGetCustomInfo(model, "PartNumber"),
                Standard = SafeGetCustomInfo(model, "Standard"),
                NominalDiameter = SafeGetCustomInfo(model, "NominalDiameter"),
                Length = SafeGetCustomInfo(model, "Length"),
                StrengthClass = SafeGetCustomInfo(model, "StrengthClass"),
                BoltType = SafeGetCustomInfo(model, "Type"),
                WasherType = SafeGetCustomInfo(model, "WasherType"),
                NutType = SafeGetCustomInfo(model, "NutType"),
                OuterDiameter = SafeGetCustomInfo(model, "OuterDiameter"),
                Thickness = SafeGetCustomInfo(model, "Thickness"),
                Height = SafeGetCustomInfo(model, "Height"),
                Description = SafeGetCustomInfo(model, "Description"),
                Family = SafeGetCustomInfo(model, "Family"),
                Material = SafeGetCustomInfo(model, "Material")
            };

            // Infer from filename (Toolbox name or our naming)
            var fileName = TryGetFileNameWithoutExtension(model);
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                InferFromFileName(fileName, values);
            }

            // Infer Family from Standard if still unknown
            if (string.IsNullOrWhiteSpace(values.Family) && !string.IsNullOrWhiteSpace(values.Standard))
            {
                values.Family = GuessFamilyFromStandard(values.Standard);
            }

            // Reasonable default if still unknown
            if (string.IsNullOrWhiteSpace(values.Family))
            {
                values.Family = "Bolt";
            }

            // Default type based on standard
            if (values.Family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(values.BoltType) &&
                !string.IsNullOrWhiteSpace(values.Standard))
            {
                values.BoltType = GuessBoltType(values.Standard);
            }

            if (values.Family.Equals("Washer", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(values.WasherType))
            {
                values.WasherType = "PlainWasher";
            }

            if (values.Family.Equals("Nut", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(values.NutType))
            {
                values.NutType = "HexNut";
            }

            // Default strength class for bolts / nuts
            if (string.IsNullOrWhiteSpace(values.StrengthClass) &&
                (values.Family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) ||
                 values.Family.Equals("Nut", StringComparison.OrdinalIgnoreCase)))
            {
                values.StrengthClass = GuessDefaultStrengthClass(values.Standard);
            }

            // Auto-suggest description if missing
            if (string.IsNullOrWhiteSpace(values.Description))
            {
                values.Description = BuildDescription(values);
            }

            return values;
        }

        public static void WriteProperties(IModelDoc2 model, FastenerInitialValues values)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));
            if (values == null) throw new ArgumentNullException(nameof(values));

            var ext = model.Extension as ModelDocExtension;
            if (ext == null) return;

            // NOTE: CustomPropertyManager is an indexer in this interop
            var pm = ext.CustomPropertyManager[""];
            if (pm == null) return;

            // Numeric values stored as text but parseable
            AddOrUpdate(pm, "PartNumber", values.PartNumber);
            AddOrUpdate(pm, "Standard", CanonicalizeStandard(values.Standard));
            AddOrUpdate(pm, "NominalDiameter", NormalizeNumeric(values.NominalDiameter));
            AddOrUpdate(pm, "Length", NormalizeNumeric(values.Length));
            AddOrUpdate(pm, "StrengthClass", values.StrengthClass);

            if (!string.IsNullOrWhiteSpace(values.BoltType))
                AddOrUpdate(pm, "Type", values.BoltType);

            if (!string.IsNullOrWhiteSpace(values.WasherType))
                AddOrUpdate(pm, "WasherType", values.WasherType);

            if (!string.IsNullOrWhiteSpace(values.NutType))
                AddOrUpdate(pm, "NutType", values.NutType);

            AddOrUpdate(pm, "OuterDiameter", NormalizeNumeric(values.OuterDiameter));
            AddOrUpdate(pm, "Thickness", NormalizeNumeric(values.Thickness));
            AddOrUpdate(pm, "Height", NormalizeNumeric(values.Height));
            AddOrUpdate(pm, "Description", values.Description);
            AddOrUpdate(pm, "Family", values.Family);
            AddOrUpdate(pm, "Material", values.Material);
        }

        public static string SafeGetCustomInfo(IModelDoc2 model, string name)
        {
            if (model == null || string.IsNullOrWhiteSpace(name))
                return null;

            try
            {
                var raw = model.GetCustomInfoValue("", name);
                return string.IsNullOrWhiteSpace(raw) ? null : raw.Trim();
            }
            catch
            {
                return null;
            }
        }

        public static string BuildDescription(FastenerInitialValues v)
        {
            if (v == null) return string.Empty;

            var family = v.Family ?? string.Empty;
            var sb = new StringBuilder();

            var std = CanonicalizeStandard(v.Standard);
            if (!string.IsNullOrWhiteSpace(std))
            {
                sb.Append(std.Trim());
                sb.Append(' ');
            }

            if (!string.IsNullOrWhiteSpace(v.NominalDiameter))
            {
                sb.Append('M');
                sb.Append(NormalizeNumeric(v.NominalDiameter));
            }

            if (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(v.Length))
            {
                sb.Append('x');
                sb.Append(NormalizeNumeric(v.Length));
            }

            if (!string.IsNullOrWhiteSpace(v.StrengthClass) &&
                (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) ||
                 family.Equals("Nut", StringComparison.OrdinalIgnoreCase)))
            {
                sb.Append(' ');
                sb.Append(v.StrengthClass.Trim());
            }

            string type = null;
            if (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
                type = v.BoltType;
            else if (family.Equals("Washer", StringComparison.OrdinalIgnoreCase))
                type = v.WasherType;
            else if (family.Equals("Nut", StringComparison.OrdinalIgnoreCase))
                type = v.NutType;

            if (!string.IsNullOrWhiteSpace(type))
            {
                sb.Append(" (");
                sb.Append(type.Trim());
                sb.Append(')');
            }

            return sb.ToString().Trim();
        }

        /// <summary>
        /// Builds the standard file name according to our convention:
        /// Bolts:  PartNumber_Standard_MdaxL-Strength
        /// Washers: PartNumber_Standard_Mda
        /// Nuts:   PartNumber_Standard_Mda-Strength
        /// </summary>
        public static string BuildStandardFileName(FastenerInitialValues v, string currentExtension)
        {
            if (v == null) throw new ArgumentNullException(nameof(v));

            if (string.IsNullOrWhiteSpace(v.PartNumber) ||
                string.IsNullOrWhiteSpace(v.Standard) ||
                string.IsNullOrWhiteSpace(v.NominalDiameter))
            {
                // Not enough information to build a standard file name
                return null;
            }

            var ext = string.IsNullOrWhiteSpace(currentExtension)
                ? ".sldprt"
                : currentExtension;

            var family = v.Family ?? "Bolt";

            var sizeToken = "M" + NormalizeNumeric(v.NominalDiameter);
            var strengthToken = string.Empty;

            if (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(v.Length))
                {
                    sizeToken += "x" + NormalizeNumeric(v.Length);
                }

                if (!string.IsNullOrWhiteSpace(v.StrengthClass))
                {
                    strengthToken = "-" + v.StrengthClass.Trim();
                }
            }
            else if (family.Equals("Nut", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(v.StrengthClass))
                {
                    strengthToken = "-" + v.StrengthClass.Trim();
                }
            }
            // Washers: M + nominal diameter only, no strength in file name.

            var std = CanonicalizeStandard(v.Standard);
            var partNumber = v.PartNumber.Trim();

            return $"{partNumber}_{std}_{sizeToken}{strengthToken}{ext}";
        }

        // ----------------- internal helpers -----------------

        private static void AddOrUpdate(ICustomPropertyManager pm, string name, string value)
        {
            if (pm == null || string.IsNullOrWhiteSpace(name))
                return;

            value ??= string.Empty;

            pm.Add3(
                name,
                (int)swCustomInfoType_e.swCustomInfoText,
                value,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
        }

        private static string TryGetFileNameWithoutExtension(IModelDoc2 model)
        {
            string path = null;
            try { path = model.GetPathName(); } catch { }

            if (string.IsNullOrWhiteSpace(path))
            {
                try { return model.GetTitle(); } catch { return null; }
            }

            try
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(path);
                return string.IsNullOrWhiteSpace(fileName) ? null : fileName;
            }
            catch
            {
                return null;
            }
        }

        private static void InferFromFileName(string fileName, FastenerInitialValues v)
        {
            if (string.IsNullOrWhiteSpace(fileName) || v == null) return;

            // 1) Our own naming format: PartNumber_Standard_M4x10-8.8
            var tokens = fileName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            bool looksLikeOurNaming =
                tokens.Length >= 2 &&
                (tokens[1].StartsWith("ISO", StringComparison.OrdinalIgnoreCase) ||
                 tokens[1].StartsWith("DIN", StringComparison.OrdinalIgnoreCase));

            if (looksLikeOurNaming)
            {
                if (tokens.Length >= 1 && string.IsNullOrWhiteSpace(v.PartNumber))
                    v.PartNumber = tokens[0];

                if (tokens.Length >= 2 && string.IsNullOrWhiteSpace(v.Standard))
                    v.Standard = CanonicalizeStandard(tokens[1]);

                if (tokens.Length >= 3)
                    ParseSizeToken(tokens[2], v);

                return;
            }

            // 2) Toolbox naming format, e.g. "ISO 4017 - M 8 x 45 - N"
            InferFromToolboxFileName(fileName, v);
        }

        // Handles Toolbox‑style names such as:
        // "ISO 4017 - M8 x 45-N"
        // "ISO 4032 - M8"
        private static void InferFromToolboxFileName(string fileName, FastenerInitialValues v)
        {
            var name = fileName.Trim();

            // Bolt pattern with M<size> x <length>
            var boltRx = new Regex(
                @"^(ISO|DIN)\s*([0-9\-]+)\s*-\s*M\s*([0-9]+(?:[.,][0-9]+)?)\s*x\s*([0-9]+(?:[.,][0-9]+)?)",
                RegexOptions.IgnoreCase);

            var m = boltRx.Match(name);
            if (m.Success)
            {
                var stdPrefix = m.Groups[1].Value.ToUpperInvariant();
                var stdNum = m.Groups[2].Value;
                var dia = m.Groups[3].Value.Replace(',', '.');
                var len = m.Groups[4].Value.Replace(',', '.');

                if (string.IsNullOrWhiteSpace(v.Standard))
                    v.Standard = stdPrefix + stdNum;

                if (string.IsNullOrWhiteSpace(v.NominalDiameter))
                    v.NominalDiameter = dia;

                if (string.IsNullOrWhiteSpace(v.Length))
                    v.Length = len;

                if (string.IsNullOrWhiteSpace(v.Family))
                    v.Family = "Bolt";

                return;
            }

            // Nuts / washers often look like "ISO 4032 - M 8" or "ISO 7089 - 8.4"
            var nutWashRx = new Regex(
                @"^(ISO|DIN)\s*([0-9\-]+)\s*-\s*([0-9M][0-9.,]*)",
                RegexOptions.IgnoreCase);

            var m2 = nutWashRx.Match(name);
            if (m2.Success)
            {
                var stdPrefix = m2.Groups[1].Value.ToUpperInvariant();
                var stdNum = m2.Groups[2].Value;
                var third = m2.Groups[3].Value.Trim();

                if (string.IsNullOrWhiteSpace(v.Standard))
                    v.Standard = stdPrefix + stdNum;

                var diaText = third;
                if (diaText.StartsWith("M", StringComparison.OrdinalIgnoreCase))
                    diaText = diaText.Substring(1);

                diaText = diaText.Replace(',', '.');

                if (string.IsNullOrWhiteSpace(v.NominalDiameter))
                    v.NominalDiameter = diaText;

                if (string.IsNullOrWhiteSpace(v.Family))
                    v.Family = GuessFamilyFromStandard(stdPrefix + stdNum);
            }
        }

        // M4x10-8.8, M4-8.8 or M4
        private static void ParseSizeToken(string token, FastenerInitialValues v)
        {
            if (string.IsNullOrWhiteSpace(token) || v == null)
                return;

            var t = token.Trim();
            int mIdx = t.IndexOf('M');
            if (mIdx < 0)
                return;

            int i = mIdx + 1;
            var dia = new StringBuilder();
            while (i < t.Length && (char.IsDigit(t[i]) || t[i] == '.' || t[i] == ','))
            {
                dia.Append(t[i]);
                i++;
            }

            if (dia.Length > 0 && string.IsNullOrWhiteSpace(v.NominalDiameter))
                v.NominalDiameter = dia.ToString();

            if (i < t.Length && (t[i] == 'x' || t[i] == 'X'))
            {
                i++;
                var len = new StringBuilder();
                while (i < t.Length && (char.IsDigit(t[i]) || t[i] == '.' || t[i] == ','))
                {
                    len.Append(t[i]);
                    i++;
                }

                if (len.Length > 0 && string.IsNullOrWhiteSpace(v.Length))
                    v.Length = len.ToString();
            }

            int dashIdx = t.IndexOf('-', Math.Max(i, 0));
            if (dashIdx >= 0 && dashIdx + 1 < t.Length && string.IsNullOrWhiteSpace(v.StrengthClass))
            {
                v.StrengthClass = t.Substring(dashIdx + 1);
            }
        }

        private static string GuessFamilyFromStandard(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return null;

            var std = CanonicalizeStandard(standard);

            // Bolts / screws
            if (std.StartsWith("ISO4014", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4017", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4762", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7045", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7046", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO2009", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO2010", StringComparison.OrdinalIgnoreCase))
            {
                return "Bolt";
            }

            // Washers
            if (std.StartsWith("ISO7089", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7090", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7092", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7093", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO7094", StringComparison.OrdinalIgnoreCase))
            {
                return "Washer";
            }

            // Nuts
            if (std.StartsWith("ISO4032", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4033", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO8673", StringComparison.OrdinalIgnoreCase))
            {
                return "Nut";
            }

            return null;
        }

        private static string GuessBoltType(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return null;

            var std = CanonicalizeStandard(standard);

            if (std.StartsWith("ISO4014", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4017", StringComparison.OrdinalIgnoreCase))
                return "HexBolt";

            if (std.StartsWith("ISO4762", StringComparison.OrdinalIgnoreCase))
                return "SocketHeadCap";

            if (std.StartsWith("ISO7045", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO2009", StringComparison.OrdinalIgnoreCase))
                return "PanHead";

            if (std.StartsWith("ISO7046", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO2010", StringComparison.OrdinalIgnoreCase))
                return "Countersunk";

            return null;
        }

        private static string GuessDefaultStrengthClass(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return "8.8";

            var std = CanonicalizeStandard(standard);

            if (std.StartsWith("ISO4014", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4017", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO4762", StringComparison.OrdinalIgnoreCase))
            {
                return "8.8";
            }

            if (std.StartsWith("ISO4032", StringComparison.OrdinalIgnoreCase) ||
                std.StartsWith("ISO8673", StringComparison.OrdinalIgnoreCase))
            {
                return "8";
            }

            return "8.8";
        }

        private static string NormalizeNumeric(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var raw = text.Trim();

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                return d.ToString("0.###", CultureInfo.InvariantCulture);

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out d))
                return d.ToString("0.###", CultureInfo.InvariantCulture);

            // Fallback: store as-is
            return raw;
        }

        private static string CanonicalizeStandard(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return string.Empty;

            // Remove spaces so "ISO 4017" -> "ISO4017"
            var s = standard.Trim().Replace(" ", string.Empty);
            return s;
        }
    }
}
