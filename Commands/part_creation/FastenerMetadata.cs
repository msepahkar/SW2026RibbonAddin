using System;
using System.Globalization;
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
    /// Helper to read/write fastener custom properties and to infer defaults
    /// from filename / ISO standard.
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

            // 1) Infer from filename (PartNumber, Standard, size / length / strength)
            var fileName = TryGetFileNameWithoutExtension(model);
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                InferFromFileName(fileName, values);
            }

            // 2) Infer Family from Standard if still unknown
            if (string.IsNullOrWhiteSpace(values.Family) && !string.IsNullOrWhiteSpace(values.Standard))
            {
                values.Family = GuessFamilyFromStandard(values.Standard);
            }

            // Reasonable default if still unknown
            if (string.IsNullOrWhiteSpace(values.Family))
            {
                values.Family = "Bolt";
            }

            // 3) Default type based on standard
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

            // 4) Default strength class for bolts / nuts
            if (string.IsNullOrWhiteSpace(values.StrengthClass) &&
                (values.Family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) ||
                 values.Family.Equals("Nut", StringComparison.OrdinalIgnoreCase)))
            {
                values.StrengthClass = GuessDefaultStrengthClass(values.Standard);
            }

            // 5) Auto-suggest description if missing
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

            var pm = ext.CustomPropertyManager[""];
            if (pm == null) return;

            // Numeric values stored as text but parseable
            AddOrUpdate(pm, "PartNumber", values.PartNumber);
            AddOrUpdate(pm, "Standard", values.Standard);
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
            var sb = new System.Text.StringBuilder();

            if (!string.IsNullOrWhiteSpace(v.Standard))
            {
                sb.Append(v.Standard.Trim());
                sb.Append(' ');
            }

            if (!string.IsNullOrWhiteSpace(v.NominalDiameter))
            {
                sb.Append('M');
                sb.Append(v.NominalDiameter.Trim());
            }

            if (family.Equals("Bolt", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(v.Length))
            {
                sb.Append('x');
                sb.Append(v.Length.Trim());
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

        // ----------------- internal helpers -----------------

        private static void AddOrUpdate(ICustomPropertyManager pm, string name, string value)
        {
            if (pm == null || string.IsNullOrWhiteSpace(name))
                return;

            value ??= string.Empty;

            // 30 = swCustomInfoType_e.swCustomInfoText
            // 1  = swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
            pm.Add3(name,
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

            // 411001-000100-00_ISO4017_M4x10-8.8
            var tokens = fileName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length >= 1 && string.IsNullOrWhiteSpace(v.PartNumber))
                v.PartNumber = tokens[0];

            if (tokens.Length >= 2 && string.IsNullOrWhiteSpace(v.Standard))
                v.Standard = tokens[1];

            if (tokens.Length >= 3)
                ParseSizeToken(tokens[2], v);
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
            var dia = new System.Text.StringBuilder();
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
                var len = new System.Text.StringBuilder();
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

            var std = standard.ToUpperInvariant();

            // Bolts / screws
            if (std.StartsWith("ISO4014") || std.StartsWith("ISO4017") ||
                std.StartsWith("ISO4762") || std.StartsWith("ISO7045") ||
                std.StartsWith("ISO7046") || std.StartsWith("ISO2009") ||
                std.StartsWith("ISO2010"))
            {
                return "Bolt";
            }

            // Washers
            if (std.StartsWith("ISO7089") || std.StartsWith("ISO7090") ||
                std.StartsWith("ISO7092") || std.StartsWith("ISO7093") ||
                std.StartsWith("ISO7094"))
            {
                return "Washer";
            }

            // Nuts
            if (std.StartsWith("ISO4032") || std.StartsWith("ISO4033") ||
                std.StartsWith("ISO8673"))
            {
                return "Nut";
            }

            return null;
        }

        private static string GuessBoltType(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return null;

            var std = standard.ToUpperInvariant();

            if (std.StartsWith("ISO4014") || std.StartsWith("ISO4017"))
                return "HexBolt";

            if (std.StartsWith("ISO4762"))
                return "SocketHeadCap";

            if (std.StartsWith("ISO7045") || std.StartsWith("ISO2009"))
                return "PanHead";

            if (std.StartsWith("ISO7046") || std.StartsWith("ISO2010"))
                return "Countersunk";

            return null;
        }

        private static string GuessDefaultStrengthClass(string standard)
        {
            if (string.IsNullOrWhiteSpace(standard))
                return "8.8";

            var std = standard.ToUpperInvariant();

            if (std.StartsWith("ISO4014") || std.StartsWith("ISO4017") ||
                std.StartsWith("ISO4762"))
            {
                return "8.8";
            }

            if (std.StartsWith("ISO4032") || std.StartsWith("ISO8673"))
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
    }
}
