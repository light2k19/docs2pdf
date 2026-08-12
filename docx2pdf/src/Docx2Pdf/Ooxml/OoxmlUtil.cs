using System;
using System.Globalization;
using System.Xml.Linq;

namespace Docx2Pdf.Ooxml
{
    /// <summary>Unit conversions and small XML helpers shared by the OOXML readers.</summary>
    internal static class OoxmlUtil
    {
        public const double TwipsPerPoint = 20.0;
        public const double EmuPerPoint = 12700.0;
        public const double PointsPerInch = 72.0;

        public static double TwipsToPoints(double twips) { return twips / TwipsPerPoint; }
        public static double EmuToPoints(double emu) { return emu / EmuPerPoint; }
        public static double EighthPointsToPoints(double eighths) { return eighths / 8.0; }
        public static double HalfPointsToPoints(double halfPoints) { return halfPoints / 2.0; }

        /// <summary>Reads an attribute as an invariant-culture double; returns null when absent/unparsable.</summary>
        public static double? Dbl(XElement el, XName name)
        {
            if (el == null)
                return null;
            var attr = el.Attribute(name);
            if (attr == null)
                return null;
            double v;
            return double.TryParse(attr.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? (double?)v : null;
        }

        public static int? Int(XElement el, XName name)
        {
            var d = Dbl(el, name);
            return d.HasValue ? (int?)(int)Math.Round(d.Value) : null;
        }

        public static string Str(XElement el, XName name)
        {
            if (el == null)
                return null;
            var attr = el.Attribute(name);
            return attr == null ? null : attr.Value;
        }

        /// <summary>Value of the <c>w:val</c> attribute of a child element.</summary>
        public static string ChildVal(XElement parent, XName child)
        {
            if (parent == null)
                return null;
            var el = parent.Element(child);
            return el == null ? null : (string)el.Attribute(Ns.W + "val");
        }

        /// <summary>
        /// OOXML toggle semantics: element absent =&gt; null, present without w:val =&gt; true,
        /// w:val of 0/false/off =&gt; false, anything else =&gt; true.
        /// </summary>
        public static bool? Toggle(XElement parent, XName child)
        {
            if (parent == null)
                return null;
            var el = parent.Element(child);
            if (el == null)
                return null;
            return ToggleValue(el);
        }

        public static bool ToggleValue(XElement el)
        {
            var val = (string)el.Attribute(Ns.W + "val");
            if (val == null)
                return true;
            switch (val.Trim().ToLowerInvariant())
            {
                case "0":
                case "false":
                case "off":
                case "none":
                    return false;
                default:
                    return true;
            }
        }

        /// <summary>Parses an OOXML "RRGGBB" colour; returns null for "auto"/empty/invalid.</summary>
        public static uint? ParseColor(string hex)
        {
            if (string.IsNullOrEmpty(hex))
                return null;
            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal))
                hex = hex.Substring(1);
            if (string.Equals(hex, "auto", StringComparison.OrdinalIgnoreCase))
                return null;
            if (hex.Length != 6)
                return null;
            uint v;
            return uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v) ? (uint?)v : null;
        }

        /// <summary>Maps a w:highlight colour keyword to RGB.</summary>
        public static uint? HighlightColor(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            switch (name.ToLowerInvariant())
            {
                case "black": return 0x000000;
                case "blue": return 0x0000FF;
                case "cyan": return 0x00FFFF;
                case "green": return 0x008000;
                case "magenta": return 0xFF00FF;
                case "red": return 0xFF0000;
                case "yellow": return 0xFFFF00;
                case "white": return 0xFFFFFF;
                case "darkblue": return 0x000080;
                case "darkcyan": return 0x008080;
                case "darkgreen": return 0x006400;
                case "darkmagenta": return 0x800080;
                case "darkred": return 0x800000;
                case "darkyellow": return 0x808000;
                case "darkgray":
                case "darkgrey": return 0x808080;
                case "lightgray":
                case "lightgrey": return 0xC0C0C0;
                case "none": return null;
                default: return null;
            }
        }

        /// <summary>
        /// Returns the effective children of a container, resolving markup-compatibility
        /// AlternateContent blocks to their first Choice (falling back to Fallback).
        /// </summary>
        public static System.Collections.Generic.IEnumerable<XElement> EffectiveElements(XElement parent)
        {
            foreach (var child in parent.Elements())
            {
                if (child.Name == Ns.Mc + "AlternateContent")
                {
                    var chosen = child.Element(Ns.Mc + "Choice") ?? child.Element(Ns.Mc + "Fallback");
                    if (chosen != null)
                    {
                        foreach (var inner in EffectiveElements(chosen))
                            yield return inner;
                    }
                    continue;
                }
                yield return child;
            }
        }
    }
}
