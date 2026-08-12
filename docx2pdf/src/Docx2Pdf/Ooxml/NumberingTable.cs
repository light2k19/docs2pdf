using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Docx2Pdf.Model;

namespace Docx2Pdf.Ooxml
{
    internal enum NumberFormat
    {
        Decimal, LowerLetter, UpperLetter, LowerRoman, UpperRoman, Bullet, None,
        DecimalZero, Ordinal, CardinalText, OrdinalText, Chicago, IdeographDigital, ChineseCounting
    }

    internal sealed class NumberingLevel
    {
        public int Level;
        public int Start = 1;
        public NumberFormat Format = NumberFormat.Decimal;
        public string LevelText = "%1.";
        public string Suffix = "tab";       // tab | space | nothing
        public bool IsLegal;
        public int? RestartAfterLevel;
        public double IndentLeftPt;
        public double HangingPt;
        public bool HasIndent;
        public CharacterFormat RunFormat;
        public string BulletFont;
        /// <summary>Paragraph style this level is bound to (w:lvl/w:pStyle).</summary>
        public string ParagraphStyle;
    }

    internal sealed class AbstractNumbering
    {
        public string Id;
        public string StyleLink;
        public string NumStyleLink;
        public readonly Dictionary<int, NumberingLevel> Levels = new Dictionary<int, NumberingLevel>();
    }

    internal sealed class NumberingInstance
    {
        public string NumId;
        public string AbstractNumId;
        public readonly Dictionary<int, NumberingLevel> Overrides = new Dictionary<int, NumberingLevel>();
        public readonly Dictionary<int, int> StartOverrides = new Dictionary<int, int>();
    }

    /// <summary>numbering.xml: abstract definitions, concrete instances and the running counter state.</summary>
    internal sealed class NumberingTable
    {
        private readonly Dictionary<string, AbstractNumbering> _abstracts = new Dictionary<string, AbstractNumbering>(StringComparer.Ordinal);
        private readonly Dictionary<string, NumberingInstance> _instances = new Dictionary<string, NumberingInstance>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _styleLinks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Counter state: key = "numId:level"
        private readonly Dictionary<string, int> _counters = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly HashSet<string> _started = new HashSet<string>(StringComparer.Ordinal);

        public static NumberingTable Parse(XDocument doc, FormatReader fmt)
        {
            var table = new NumberingTable();
            if (doc == null || doc.Root == null)
                return table;

            foreach (var el in doc.Root.Elements(Ns.W + "abstractNum"))
            {
                var abs = new AbstractNumbering { Id = OoxmlUtil.Str(el, Ns.W + "abstractNumId") };
                if (abs.Id == null)
                    continue;
                abs.StyleLink = OoxmlUtil.ChildVal(el, Ns.W + "styleLink");
                abs.NumStyleLink = OoxmlUtil.ChildVal(el, Ns.W + "numStyleLink");
                foreach (var lvlEl in el.Elements(Ns.W + "lvl"))
                {
                    var lvl = ParseLevel(lvlEl, fmt);
                    if (lvl != null)
                        abs.Levels[lvl.Level] = lvl;
                }
                table._abstracts[abs.Id] = abs;
                if (!string.IsNullOrEmpty(abs.StyleLink))
                    table._styleLinks[abs.StyleLink] = abs.Id;
            }

            foreach (var el in doc.Root.Elements(Ns.W + "num"))
            {
                var inst = new NumberingInstance
                {
                    NumId = OoxmlUtil.Str(el, Ns.W + "numId"),
                    AbstractNumId = OoxmlUtil.ChildVal(el, Ns.W + "abstractNumId"),
                };
                if (inst.NumId == null)
                    continue;
                foreach (var ov in el.Elements(Ns.W + "lvlOverride"))
                {
                    int ilvl;
                    if (!int.TryParse(OoxmlUtil.Str(ov, Ns.W + "ilvl") ?? "0", out ilvl))
                        continue;
                    var startOverride = OoxmlUtil.ChildVal(ov, Ns.W + "startOverride");
                    int so;
                    if (startOverride != null && int.TryParse(startOverride, out so))
                        inst.StartOverrides[ilvl] = so;
                    var lvlEl = ov.Element(Ns.W + "lvl");
                    if (lvlEl != null)
                    {
                        var lvl = ParseLevel(lvlEl, fmt);
                        if (lvl != null)
                        {
                            lvl.Level = ilvl;
                            inst.Overrides[ilvl] = lvl;
                        }
                    }
                }
                table._instances[inst.NumId] = inst;
            }

            return table;
        }

        private static NumberingLevel ParseLevel(XElement el, FormatReader fmt)
        {
            int ilvl;
            if (!int.TryParse(OoxmlUtil.Str(el, Ns.W + "ilvl") ?? "0", out ilvl))
                ilvl = 0;

            var lvl = new NumberingLevel { Level = ilvl };

            var start = OoxmlUtil.ChildVal(el, Ns.W + "start");
            int s;
            if (start != null && int.TryParse(start, out s))
                lvl.Start = s;

            lvl.Format = ParseFormat(OoxmlUtil.ChildVal(el, Ns.W + "numFmt"));
            var text = OoxmlUtil.ChildVal(el, Ns.W + "lvlText");
            if (text != null)
                lvl.LevelText = text;
            var suff = OoxmlUtil.ChildVal(el, Ns.W + "suff");
            if (suff != null)
                lvl.Suffix = suff.ToLowerInvariant();
            lvl.IsLegal = OoxmlUtil.Toggle(el, Ns.W + "isLgl") == true;
            var restart = OoxmlUtil.ChildVal(el, Ns.W + "lvlRestart");
            int r;
            if (restart != null && int.TryParse(restart, out r))
                lvl.RestartAfterLevel = r;

            lvl.ParagraphStyle = OoxmlUtil.ChildVal(el, Ns.W + "pStyle");

            var pPr = el.Element(Ns.W + "pPr");
            if (pPr != null)
            {
                var ind = pPr.Element(Ns.W + "ind");
                if (ind != null)
                {
                    var left = OoxmlUtil.Dbl(ind, Ns.W + "left") ?? OoxmlUtil.Dbl(ind, Ns.W + "start");
                    var hanging = OoxmlUtil.Dbl(ind, Ns.W + "hanging");
                    if (left.HasValue)
                    {
                        lvl.IndentLeftPt = OoxmlUtil.TwipsToPoints(left.Value);
                        lvl.HasIndent = true;
                    }
                    if (hanging.HasValue)
                    {
                        lvl.HangingPt = OoxmlUtil.TwipsToPoints(hanging.Value);
                        lvl.HasIndent = true;
                    }
                }
            }

            var rPr = el.Element(Ns.W + "rPr");
            if (rPr != null)
            {
                lvl.RunFormat = fmt.ReadRunFormat(rPr);
                var rFonts = rPr.Element(Ns.W + "rFonts");
                if (rFonts != null)
                    lvl.BulletFont = OoxmlUtil.Str(rFonts, Ns.W + "ascii") ?? OoxmlUtil.Str(rFonts, Ns.W + "hAnsi");
            }

            return lvl;
        }

        private static NumberFormat ParseFormat(string val)
        {
            if (string.IsNullOrEmpty(val))
                return NumberFormat.Decimal;
            switch (val.ToLowerInvariant())
            {
                case "bullet": return NumberFormat.Bullet;
                case "none": return NumberFormat.None;
                case "lowerletter": return NumberFormat.LowerLetter;
                case "upperletter": return NumberFormat.UpperLetter;
                case "lowerroman": return NumberFormat.LowerRoman;
                case "upperroman": return NumberFormat.UpperRoman;
                case "decimalzero": return NumberFormat.DecimalZero;
                case "ordinal": return NumberFormat.Ordinal;
                case "cardinaltext": return NumberFormat.CardinalText;
                case "ordinaltext": return NumberFormat.OrdinalText;
                case "chicago": return NumberFormat.Chicago;
                case "ideographdigital":
                case "japanesecounting":
                case "taiwanesecounting": return NumberFormat.IdeographDigital;
                case "chinesecounting": return NumberFormat.ChineseCounting;
                default: return NumberFormat.Decimal;
            }
        }

        public NumberingLevel GetLevel(string numId, int ilvl)
        {
            NumberingInstance inst;
            if (numId == null || !_instances.TryGetValue(numId, out inst))
                return null;

            NumberingLevel ovr;
            if (inst.Overrides.TryGetValue(ilvl, out ovr))
                return ovr;

            var abs = ResolveAbstract(inst.AbstractNumId, 0);
            if (abs == null)
                return null;

            NumberingLevel lvl;
            if (abs.Levels.TryGetValue(ilvl, out lvl))
                return lvl;

            // Fall back to the deepest defined level.
            NumberingLevel best = null;
            foreach (var kv in abs.Levels)
            {
                if (best == null || Math.Abs(kv.Key - ilvl) < Math.Abs(best.Level - ilvl))
                    best = kv.Value;
            }
            return best;
        }

        private AbstractNumbering ResolveAbstract(string abstractId, int depth)
        {
            if (abstractId == null || depth > 8)
                return null;
            AbstractNumbering abs;
            if (!_abstracts.TryGetValue(abstractId, out abs))
                return null;

            if (abs.Levels.Count == 0 && !string.IsNullOrEmpty(abs.NumStyleLink))
            {
                string linked;
                if (_styleLinks.TryGetValue(abs.NumStyleLink, out linked))
                    return ResolveAbstract(linked, depth + 1);
            }
            return abs;
        }

        /// <summary>
        /// Finds the level a paragraph style is bound to through w:lvl/w:pStyle.
        /// Multi-level heading numbering relies on this: the heading styles carry no explicit level.
        /// </summary>
        public bool TryGetLevelForStyle(string numId, string styleId, out int ilvl)
        {
            ilvl = 0;
            if (string.IsNullOrEmpty(styleId))
                return false;

            NumberingInstance inst;
            if (numId == null || !_instances.TryGetValue(numId, out inst))
                return false;

            foreach (var kv in inst.Overrides)
            {
                if (string.Equals(kv.Value.ParagraphStyle, styleId, StringComparison.OrdinalIgnoreCase))
                {
                    ilvl = kv.Key;
                    return true;
                }
            }

            var abs = ResolveAbstract(inst.AbstractNumId, 0);
            if (abs == null)
                return false;

            foreach (var kv in abs.Levels)
            {
                if (string.Equals(kv.Value.ParagraphStyle, styleId, StringComparison.OrdinalIgnoreCase))
                {
                    ilvl = kv.Key;
                    return true;
                }
            }
            return false;
        }

        public bool HasInstance(string numId)
        {
            return numId != null && _instances.ContainsKey(numId);
        }

        /// <summary>
        /// Advances the counter for (numId, ilvl), resets deeper levels and renders the label text.
        /// </summary>
        public string NextLabel(string numId, int ilvl, out NumberingLevel level)
        {
            level = GetLevel(numId, ilvl);
            if (level == null)
                return null;
            if (level.Format == NumberFormat.None)
                return string.Empty;

            NumberingInstance inst;
            _instances.TryGetValue(numId, out inst);

            var key = numId + ":" + ilvl;
            int current;
            if (!_counters.TryGetValue(key, out current) || !_started.Contains(key))
            {
                var start = level.Start;
                int so;
                if (inst != null && inst.StartOverrides.TryGetValue(ilvl, out so))
                    start = so;
                current = start;
                _started.Add(key);
            }
            else
            {
                current++;
            }
            _counters[key] = current;

            // Deeper levels restart.
            for (var deeper = ilvl + 1; deeper <= 8; deeper++)
            {
                var dkey = numId + ":" + deeper;
                _started.Remove(dkey);
                _counters.Remove(dkey);
            }

            return RenderLabel(numId, ilvl, level);
        }

        private string RenderLabel(string numId, int ilvl, NumberingLevel level)
        {
            if (level.Format == NumberFormat.Bullet)
                return level.LevelText;

            var sb = new StringBuilder();
            var text = level.LevelText ?? string.Empty;
            for (var i = 0; i < text.Length; i++)
            {
                if (text[i] == '%' && i + 1 < text.Length && char.IsDigit(text[i + 1]))
                {
                    var refLevel = text[i + 1] - '1';
                    i++;
                    int value;
                    _counters.TryGetValue(numId + ":" + refLevel, out value);
                    if (value == 0)
                    {
                        var refDef = GetLevel(numId, refLevel);
                        value = refDef != null ? refDef.Start : 1;
                    }
                    var fmt = refLevel == ilvl ? level.Format : (GetLevel(numId, refLevel) ?? level).Format;
                    if (level.IsLegal && refLevel != ilvl)
                        fmt = NumberFormat.Decimal;
                    sb.Append(FormatNumber(value, fmt));
                }
                else
                {
                    sb.Append(text[i]);
                }
            }
            return sb.ToString();
        }

        public static string FormatNumber(int value, NumberFormat format)
        {
            switch (format)
            {
                case NumberFormat.Decimal:
                    return value.ToString(CultureInfo.InvariantCulture);
                case NumberFormat.DecimalZero:
                    return value < 10 ? "0" + value.ToString(CultureInfo.InvariantCulture) : value.ToString(CultureInfo.InvariantCulture);
                case NumberFormat.LowerLetter:
                    return AlphaNumber(value, false);
                case NumberFormat.UpperLetter:
                    return AlphaNumber(value, true);
                case NumberFormat.LowerRoman:
                    return RomanNumber(value).ToLowerInvariant();
                case NumberFormat.UpperRoman:
                    return RomanNumber(value);
                case NumberFormat.Ordinal:
                    return Ordinal(value);
                case NumberFormat.IdeographDigital:
                case NumberFormat.ChineseCounting:
                    return ChineseNumber(value);
                case NumberFormat.None:
                    return string.Empty;
                default:
                    return value.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static string AlphaNumber(int value, bool upper)
        {
            if (value <= 0)
                return string.Empty;
            // Word repeats the letter: 1=a .. 26=z, 27=aa, 28=bb ...
            var index = (value - 1) % 26;
            var repeat = (value - 1) / 26 + 1;
            var c = (char)((upper ? 'A' : 'a') + index);
            return new string(c, repeat);
        }

        private static string RomanNumber(int value)
        {
            if (value <= 0 || value > 3999)
                return value.ToString(CultureInfo.InvariantCulture);
            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            var sb = new StringBuilder();
            for (var i = 0; i < values.Length; i++)
            {
                while (value >= values[i])
                {
                    sb.Append(symbols[i]);
                    value -= values[i];
                }
            }
            return sb.ToString();
        }

        private static string Ordinal(int value)
        {
            var n = value.ToString(CultureInfo.InvariantCulture);
            var mod100 = value % 100;
            if (mod100 >= 11 && mod100 <= 13)
                return n + "th";
            switch (value % 10)
            {
                case 1: return n + "st";
                case 2: return n + "nd";
                case 3: return n + "rd";
                default: return n + "th";
            }
        }

        private static string ChineseNumber(int value)
        {
            const string digits = "〇一二三四五六七八九";
            if (value <= 0)
                return string.Empty;
            if (value < 10)
                return digits[value].ToString();
            if (value < 20)
                return "十" + (value % 10 == 0 ? string.Empty : digits[value % 10].ToString());
            if (value < 100)
                return digits[value / 10] + "十" + (value % 10 == 0 ? string.Empty : digits[value % 10].ToString());
            return value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Clears counter state (used between rendering passes).</summary>
        public void ResetCounters()
        {
            _counters.Clear();
            _started.Clear();
        }
    }
}
