using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Docx2Pdf.Model;

namespace Docx2Pdf.Ooxml
{
    /// <summary>Translates <c>w:pPr</c> / <c>w:rPr</c> property bags into model formatting objects.</summary>
    internal sealed class FormatReader
    {
        private readonly ThemeFonts _theme;

        public FormatReader(ThemeFonts theme)
        {
            _theme = theme ?? new ThemeFonts();
        }

        public CharacterFormat ReadRunFormat(XElement rPr)
        {
            var f = new CharacterFormat();
            if (rPr == null)
                return f;

            var rFonts = rPr.Element(Ns.W + "rFonts");
            if (rFonts != null)
            {
                // Latin text takes w:ascii / w:hAnsi only. The complex-script font (w:cs) must
                // not leak in: a style that sets only eastAsia/cs (like this template's Normal)
                // leaves the Latin font inherited, and Word renders it with the theme font.
                f.FontFamily = ResolveThemeFont(OoxmlUtil.Str(rFonts, Ns.W + "ascii"), OoxmlUtil.Str(rFonts, Ns.W + "asciiTheme"))
                            ?? ResolveThemeFont(OoxmlUtil.Str(rFonts, Ns.W + "hAnsi"), OoxmlUtil.Str(rFonts, Ns.W + "hAnsiTheme"));
                f.EastAsianFontFamily = ResolveThemeFont(OoxmlUtil.Str(rFonts, Ns.W + "eastAsia"), OoxmlUtil.Str(rFonts, Ns.W + "eastAsiaTheme"));
            }

            f.Bold = OoxmlUtil.Toggle(rPr, Ns.W + "b");
            f.Italic = OoxmlUtil.Toggle(rPr, Ns.W + "i");
            f.Strike = OoxmlUtil.Toggle(rPr, Ns.W + "strike");
            if (OoxmlUtil.Toggle(rPr, Ns.W + "dstrike") == true)
                f.Strike = true;
            f.AllCaps = OoxmlUtil.Toggle(rPr, Ns.W + "caps");
            f.SmallCaps = OoxmlUtil.Toggle(rPr, Ns.W + "smallCaps");
            var vanish = OoxmlUtil.Toggle(rPr, Ns.W + "vanish");
            if (vanish.HasValue)
                f.Hidden = vanish;

            var u = rPr.Element(Ns.W + "u");
            if (u != null)
            {
                f.Underline = ParseUnderline(OoxmlUtil.Str(u, Ns.W + "val"));
                f.UnderlineColor = OoxmlUtil.ParseColor(OoxmlUtil.Str(u, Ns.W + "color"));
            }

            var sz = rPr.Element(Ns.W + "sz");
            var szVal = OoxmlUtil.Dbl(sz, Ns.W + "val");
            if (szVal.HasValue && szVal.Value > 0)
                f.SizePt = OoxmlUtil.HalfPointsToPoints(szVal.Value);

            var color = rPr.Element(Ns.W + "color");
            if (color != null)
                f.Color = OoxmlUtil.ParseColor(OoxmlUtil.Str(color, Ns.W + "val"));

            var highlight = rPr.Element(Ns.W + "highlight");
            if (highlight != null)
                f.Highlight = OoxmlUtil.HighlightColor(OoxmlUtil.Str(highlight, Ns.W + "val"));

            var shd = rPr.Element(Ns.W + "shd");
            if (shd != null)
                f.Shading = ReadShadingFill(shd);

            var vertAlign = OoxmlUtil.ChildVal(rPr, Ns.W + "vertAlign");
            if (vertAlign != null)
            {
                if (string.Equals(vertAlign, "superscript", StringComparison.OrdinalIgnoreCase))
                    f.VertAlign = VerticalTextAlignment.Superscript;
                else if (string.Equals(vertAlign, "subscript", StringComparison.OrdinalIgnoreCase))
                    f.VertAlign = VerticalTextAlignment.Subscript;
                else
                    f.VertAlign = VerticalTextAlignment.Baseline;
            }

            var spacing = rPr.Element(Ns.W + "spacing");
            var spacingVal = OoxmlUtil.Dbl(spacing, Ns.W + "val");
            if (spacingVal.HasValue)
                f.CharacterSpacingPt = OoxmlUtil.TwipsToPoints(spacingVal.Value);

            return f;
        }

        public ParagraphFormat ReadParagraphFormat(XElement pPr)
        {
            var f = new ParagraphFormat();
            if (pPr == null)
                return f;

            var jc = OoxmlUtil.ChildVal(pPr, Ns.W + "jc");
            if (jc != null)
            {
                switch (jc.ToLowerInvariant())
                {
                    case "center": f.Alignment = TextAlignment.Center; break;
                    case "right":
                    case "end": f.Alignment = TextAlignment.Right; break;
                    case "both":
                    case "distribute":
                    case "justify": f.Alignment = TextAlignment.Justify; break;
                    default: f.Alignment = TextAlignment.Left; break;
                }
            }

            var ind = pPr.Element(Ns.W + "ind");
            if (ind != null)
            {
                var left = OoxmlUtil.Dbl(ind, Ns.W + "left") ?? OoxmlUtil.Dbl(ind, Ns.W + "start");
                if (left.HasValue) f.IndentLeftPt = OoxmlUtil.TwipsToPoints(left.Value);
                var right = OoxmlUtil.Dbl(ind, Ns.W + "right") ?? OoxmlUtil.Dbl(ind, Ns.W + "end");
                if (right.HasValue) f.IndentRightPt = OoxmlUtil.TwipsToPoints(right.Value);

                var hanging = OoxmlUtil.Dbl(ind, Ns.W + "hanging");
                var firstLine = OoxmlUtil.Dbl(ind, Ns.W + "firstLine");
                if (hanging.HasValue)
                    f.IndentFirstLinePt = -OoxmlUtil.TwipsToPoints(hanging.Value);
                else if (firstLine.HasValue)
                    f.IndentFirstLinePt = OoxmlUtil.TwipsToPoints(firstLine.Value);
            }

            var spacing = pPr.Element(Ns.W + "spacing");
            if (spacing != null)
            {
                // beforeAutospacing/afterAutospacing marks HTML "auto" spacing; the stated
                // value is kept here and the layout decides its effect from context
                // (measured against Word: collapsed at body level, kept inside table cells).
                var beforeAuto = OoxmlUtil.Str(spacing, Ns.W + "beforeAutospacing");
                if (beforeAuto != null)
                    f.AutoSpaceBefore = IsOn(beforeAuto);
                var afterAuto = OoxmlUtil.Str(spacing, Ns.W + "afterAutospacing");
                if (afterAuto != null)
                    f.AutoSpaceAfter = IsOn(afterAuto);

                var before = OoxmlUtil.Dbl(spacing, Ns.W + "before");
                if (before.HasValue) f.SpaceBeforePt = OoxmlUtil.TwipsToPoints(before.Value);
                var after = OoxmlUtil.Dbl(spacing, Ns.W + "after");
                if (after.HasValue) f.SpaceAfterPt = OoxmlUtil.TwipsToPoints(after.Value);

                var line = OoxmlUtil.Dbl(spacing, Ns.W + "line");
                if (line.HasValue)
                {
                    var rule = (OoxmlUtil.Str(spacing, Ns.W + "lineRule") ?? "auto").ToLowerInvariant();
                    if (rule == "exact")
                    {
                        f.LineSpacingRule = Model.LineSpacingRule.Exact;
                        f.LineSpacing = OoxmlUtil.TwipsToPoints(line.Value);
                    }
                    else if (rule == "atleast")
                    {
                        f.LineSpacingRule = Model.LineSpacingRule.AtLeast;
                        f.LineSpacing = OoxmlUtil.TwipsToPoints(line.Value);
                    }
                    else
                    {
                        f.LineSpacingRule = Model.LineSpacingRule.Auto;
                        f.LineSpacing = line.Value / 240.0;
                    }
                }
            }

            var contextual = OoxmlUtil.Toggle(pPr, Ns.W + "contextualSpacing");
            if (contextual.HasValue) f.ContextualSpacing = contextual;
            var keepNext = OoxmlUtil.Toggle(pPr, Ns.W + "keepNext");
            if (keepNext.HasValue) f.KeepNext = keepNext;
            var keepLines = OoxmlUtil.Toggle(pPr, Ns.W + "keepLines");
            if (keepLines.HasValue) f.KeepLines = keepLines;
            var pageBreak = OoxmlUtil.Toggle(pPr, Ns.W + "pageBreakBefore");
            if (pageBreak.HasValue) f.PageBreakBefore = pageBreak;
            var bidi = OoxmlUtil.Toggle(pPr, Ns.W + "bidi");
            if (bidi.HasValue) f.Bidi = bidi;
            var widowControl = OoxmlUtil.Toggle(pPr, Ns.W + "widowControl");
            if (widowControl.HasValue) f.WidowControl = widowControl;

            var outline = OoxmlUtil.ChildVal(pPr, Ns.W + "outlineLvl");
            int lvl;
            if (outline != null && int.TryParse(outline, out lvl))
                f.OutlineLevel = lvl;

            var shd = pPr.Element(Ns.W + "shd");
            if (shd != null)
                f.Shading = ReadShadingFill(shd);

            var pBdr = pPr.Element(Ns.W + "pBdr");
            if (pBdr != null)
                f.Borders = ReadBorders(pBdr);

            var tabs = pPr.Element(Ns.W + "tabs");
            if (tabs != null)
            {
                var list = new List<TabStop>();
                foreach (var tab in tabs.Elements(Ns.W + "tab"))
                {
                    var pos = OoxmlUtil.Dbl(tab, Ns.W + "pos");
                    if (!pos.HasValue)
                        continue;
                    var stop = new TabStop { PositionPt = OoxmlUtil.TwipsToPoints(pos.Value) };
                    var val = (OoxmlUtil.Str(tab, Ns.W + "val") ?? "left").ToLowerInvariant();
                    switch (val)
                    {
                        case "center": stop.Alignment = TabAlignment.Center; break;
                        case "right":
                        case "end": stop.Alignment = TabAlignment.Right; break;
                        case "decimal": stop.Alignment = TabAlignment.Decimal; break;
                        case "bar": stop.Alignment = TabAlignment.Bar; break;
                        case "clear": stop.Alignment = TabAlignment.Clear; break;
                        default: stop.Alignment = TabAlignment.Left; break;
                    }
                    var leader = (OoxmlUtil.Str(tab, Ns.W + "leader") ?? "none").ToLowerInvariant();
                    switch (leader)
                    {
                        case "dot": stop.Leader = TabLeader.Dot; break;
                        case "hyphen": stop.Leader = TabLeader.Hyphen; break;
                        case "underscore":
                        case "heavy": stop.Leader = TabLeader.Underscore; break;
                        default: stop.Leader = TabLeader.None; break;
                    }
                    list.Add(stop);
                }
                if (list.Count > 0)
                    f.Tabs = list;
            }

            return f;
        }

        private static bool IsOn(string value)
        {
            return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        public static uint? ReadShadingFill(XElement shd)
        {
            if (shd == null)
                return null;
            var pattern = (OoxmlUtil.Str(shd, Ns.W + "val") ?? "clear").ToLowerInvariant();
            if (pattern == "nil" || pattern == "none")
                return null;

            var fill = OoxmlUtil.ParseColor(OoxmlUtil.Str(shd, Ns.W + "fill"));
            var color = OoxmlUtil.ParseColor(OoxmlUtil.Str(shd, Ns.W + "color"));

            if (pattern == "clear")
                return fill;

            // Percentage patterns (pct10 ... pct90) blend the pattern colour over the fill.
            double pct = 0.5;
            if (pattern.StartsWith("pct", StringComparison.Ordinal))
            {
                double p;
                if (double.TryParse(pattern.Substring(3), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out p))
                    pct = p / 100.0;
            }
            // "auto" resolves to a black pattern over a white fill: pct25 with both colours
            // auto is Word's standard 25% grey table shading.
            var bg = fill ?? 0xFFFFFFu;
            var fg = color ?? 0x000000u;
            return Blend(bg, fg, pct);
        }

        private static uint Blend(uint bg, uint fg, double f)
        {
            f = f < 0 ? 0 : (f > 1 ? 1 : f);
            uint r = (uint)Math.Round(((bg >> 16) & 0xFF) * (1 - f) + ((fg >> 16) & 0xFF) * f);
            uint g = (uint)Math.Round(((bg >> 8) & 0xFF) * (1 - f) + ((fg >> 8) & 0xFF) * f);
            uint b = (uint)Math.Round((bg & 0xFF) * (1 - f) + (fg & 0xFF) * f);
            return (r << 16) | (g << 8) | b;
        }

        public static Borders ReadBorders(XElement parent)
        {
            if (parent == null)
                return null;
            var b = new Borders
            {
                Top = ReadBorder(parent.Element(Ns.W + "top")),
                Bottom = ReadBorder(parent.Element(Ns.W + "bottom")),
                Left = ReadBorder(parent.Element(Ns.W + "left") ?? parent.Element(Ns.W + "start")),
                Right = ReadBorder(parent.Element(Ns.W + "right") ?? parent.Element(Ns.W + "end")),
                InsideH = ReadBorder(parent.Element(Ns.W + "insideH")),
                InsideV = ReadBorder(parent.Element(Ns.W + "insideV")),
            };
            return b;
        }

        public static Border ReadBorder(XElement el)
        {
            if (el == null)
                return null;
            var val = (OoxmlUtil.Str(el, Ns.W + "val") ?? "none").ToLowerInvariant();
            var border = new Border();
            switch (val)
            {
                case "none":
                case "nil":
                    border.Style = BorderStyle.None;
                    break;
                case "double":
                case "doublewave":
                case "triple":
                    border.Style = BorderStyle.Double;
                    break;
                case "dotted":
                case "dotdash":
                case "dotdotdash":
                    border.Style = BorderStyle.Dotted;
                    break;
                case "dashed":
                case "dashsmallgap":
                case "dashdotstroked":
                    border.Style = BorderStyle.Dashed;
                    break;
                case "thick":
                case "thickthinsmallgap":
                case "thinthicksmallgap":
                    border.Style = BorderStyle.Thick;
                    break;
                default:
                    border.Style = BorderStyle.Single;
                    break;
            }

            var sz = OoxmlUtil.Dbl(el, Ns.W + "sz");
            border.WidthPt = sz.HasValue ? Math.Max(0.25, OoxmlUtil.EighthPointsToPoints(sz.Value)) : 0.5;
            border.Color = OoxmlUtil.ParseColor(OoxmlUtil.Str(el, Ns.W + "color")) ?? 0x000000u;
            var space = OoxmlUtil.Dbl(el, Ns.W + "space");
            border.SpacePt = space ?? 0;
            return border;
        }

        private string ResolveThemeFont(string explicitName, string themeName)
        {
            if (!string.IsNullOrEmpty(themeName))
            {
                var resolved = _theme.Resolve(themeName);
                if (!string.IsNullOrEmpty(resolved))
                    return resolved;
            }
            if (string.IsNullOrEmpty(explicitName))
                return null;
            if (explicitName.StartsWith("+", StringComparison.Ordinal))
                return _theme.Resolve(explicitName.Substring(1));
            return explicitName;
        }

        public static UnderlineStyle ParseUnderline(string val)
        {
            if (string.IsNullOrEmpty(val))
                return UnderlineStyle.Single;
            switch (val.ToLowerInvariant())
            {
                case "none": return UnderlineStyle.None;
                case "double":
                case "dotteddouble":
                case "dashdotdotheavy": return UnderlineStyle.Double;
                case "thick":
                case "wavyheavy":
                case "dottedheavy":
                case "dashedheavy": return UnderlineStyle.Thick;
                case "dotted": return UnderlineStyle.Dotted;
                case "dash":
                case "dashed":
                case "dashlong":
                case "dotdash":
                case "dotdotdash": return UnderlineStyle.Dashed;
                case "wave":
                case "wavydouble": return UnderlineStyle.Wave;
                default: return UnderlineStyle.Single;
            }
        }
    }

    /// <summary>Theme font scheme lookup (<c>+mn-lt</c>, <c>minorHAnsi</c>, ...).</summary>
    internal sealed class ThemeFonts
    {
        public string MajorLatin = "Calibri Light";
        public string MinorLatin = "Calibri";
        public string MajorEastAsian;
        public string MinorEastAsian;

        public string Resolve(string token)
        {
            if (string.IsNullOrEmpty(token))
                return null;
            switch (token.ToLowerInvariant())
            {
                case "mj-lt":
                case "majorhansi":
                case "majorascii":
                case "majorbidi":
                    return MajorLatin;
                case "mn-lt":
                case "minorhansi":
                case "minorascii":
                case "minorbidi":
                    return MinorLatin;
                case "mj-ea":
                case "majoreastasia":
                    return string.IsNullOrEmpty(MajorEastAsian) ? MajorLatin : MajorEastAsian;
                case "mn-ea":
                case "minoreastasia":
                    return string.IsNullOrEmpty(MinorEastAsian) ? MinorLatin : MinorEastAsian;
                default:
                    return null;
            }
        }

        public static ThemeFonts Parse(XDocument theme)
        {
            var t = new ThemeFonts();
            if (theme == null || theme.Root == null)
                return t;
            var scheme = theme.Root.Descendants(Ns.A + "fontScheme").FirstOrDefaultSafe();
            if (scheme == null)
                return t;

            var major = scheme.Element(Ns.A + "majorFont");
            var minor = scheme.Element(Ns.A + "minorFont");
            if (major != null)
            {
                t.MajorLatin = Typeface(major, "latin") ?? t.MajorLatin;
                t.MajorEastAsian = Typeface(major, "ea");
            }
            if (minor != null)
            {
                t.MinorLatin = Typeface(minor, "latin") ?? t.MinorLatin;
                t.MinorEastAsian = Typeface(minor, "ea");
            }
            return t;
        }

        private static string Typeface(XElement font, string local)
        {
            var el = font.Element(Ns.A + local);
            if (el == null)
                return null;
            var tf = (string)el.Attribute("typeface");
            return string.IsNullOrEmpty(tf) ? null : tf;
        }
    }

    internal static class LinqSafe
    {
        public static XElement FirstOrDefaultSafe(this System.Collections.Generic.IEnumerable<XElement> src)
        {
            foreach (var e in src)
                return e;
            return null;
        }
    }
}
