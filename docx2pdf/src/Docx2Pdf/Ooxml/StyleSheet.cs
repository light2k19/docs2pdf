using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docx2Pdf.Model;

namespace Docx2Pdf.Ooxml
{
    internal sealed class WordStyle
    {
        public string Id;
        public string Name;
        public string Type;           // paragraph | character | table | numbering
        public string BasedOn;
        public bool IsDefault;
        public XElement PPr;
        public XElement RPr;
        public XElement TblPr;
        public XElement TcPr;
        /// <summary>Conditional-region formatting (w:tblStylePr) of a table style.</summary>
        public List<XElement> TblStylePrs;
        public string NumId;
        public int? NumLevel;
        public int? OutlineLevel;
    }

    /// <summary>styles.xml: document defaults plus the style hierarchy with basedOn resolution.</summary>
    internal sealed class StyleSheet
    {
        private readonly Dictionary<string, WordStyle> _styles = new Dictionary<string, WordStyle>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, WordStyle> _byName = new Dictionary<string, WordStyle>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ParagraphFormat> _pCache = new Dictionary<string, ParagraphFormat>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, CharacterFormat> _rCache = new Dictionary<string, CharacterFormat>(StringComparer.OrdinalIgnoreCase);
        private readonly FormatReader _fmt;

        public ParagraphFormat DefaultParagraphFormat = ParagraphFormat.Default();
        public CharacterFormat DefaultCharacterFormat = CharacterFormat.Default();
        public string DefaultParagraphStyleId;
        public string DefaultTableStyleId;

        public StyleSheet(FormatReader fmt)
        {
            _fmt = fmt;
        }

        public static StyleSheet Parse(XDocument doc, FormatReader fmt)
        {
            var sheet = new StyleSheet(fmt);
            if (doc == null || doc.Root == null)
                return sheet;

            var docDefaults = doc.Root.Element(Ns.W + "docDefaults");
            if (docDefaults != null)
            {
                var rPrDefault = docDefaults.Element(Ns.W + "rPrDefault");
                if (rPrDefault != null)
                {
                    var rPr = rPrDefault.Element(Ns.W + "rPr");
                    var cf = fmt.ReadRunFormat(rPr);
                    var baseCf = CharacterFormat.Default();
                    baseCf.ApplyOver(cf);
                    sheet.DefaultCharacterFormat = baseCf;
                }
                var pPrDefault = docDefaults.Element(Ns.W + "pPrDefault");
                if (pPrDefault != null)
                {
                    var pPr = pPrDefault.Element(Ns.W + "pPr");
                    var pf = fmt.ReadParagraphFormat(pPr);
                    var basePf = ParagraphFormat.Default();
                    basePf.ApplyOver(pf);
                    sheet.DefaultParagraphFormat = basePf;
                }
            }

            foreach (var el in doc.Root.Elements(Ns.W + "style"))
            {
                var style = new WordStyle
                {
                    Id = OoxmlUtil.Str(el, Ns.W + "styleId"),
                    Type = (OoxmlUtil.Str(el, Ns.W + "type") ?? "paragraph").ToLowerInvariant(),
                    IsDefault = OoxmlUtil.Str(el, Ns.W + "default") == "1"
                                || string.Equals(OoxmlUtil.Str(el, Ns.W + "default"), "true", StringComparison.OrdinalIgnoreCase),
                    Name = OoxmlUtil.ChildVal(el, Ns.W + "name"),
                    BasedOn = OoxmlUtil.ChildVal(el, Ns.W + "basedOn"),
                    PPr = el.Element(Ns.W + "pPr"),
                    RPr = el.Element(Ns.W + "rPr"),
                    TblPr = el.Element(Ns.W + "tblPr"),
                    TcPr = el.Element(Ns.W + "tcPr"),
                    TblStylePrs = el.Elements(Ns.W + "tblStylePr").ToList(),
                };
                if (string.IsNullOrEmpty(style.Id))
                    continue;

                if (style.PPr != null)
                {
                    var numPr = style.PPr.Element(Ns.W + "numPr");
                    if (numPr != null)
                    {
                        style.NumId = OoxmlUtil.ChildVal(numPr, Ns.W + "numId");
                        var ilvl = OoxmlUtil.ChildVal(numPr, Ns.W + "ilvl");
                        int lvl;
                        if (ilvl != null && int.TryParse(ilvl, out lvl))
                            style.NumLevel = lvl;
                    }
                    var outline = OoxmlUtil.ChildVal(style.PPr, Ns.W + "outlineLvl");
                    int o;
                    if (outline != null && int.TryParse(outline, out o))
                        style.OutlineLevel = o;
                }

                sheet._styles[style.Id] = style;
                if (!string.IsNullOrEmpty(style.Name) && !sheet._byName.ContainsKey(style.Name))
                    sheet._byName[style.Name] = style;

                if (style.IsDefault && style.Type == "paragraph" && sheet.DefaultParagraphStyleId == null)
                    sheet.DefaultParagraphStyleId = style.Id;
                if (style.IsDefault && style.Type == "table" && sheet.DefaultTableStyleId == null)
                    sheet.DefaultTableStyleId = style.Id;
            }

            return sheet;
        }

        public WordStyle Get(string styleId)
        {
            if (string.IsNullOrEmpty(styleId))
                return null;
            WordStyle s;
            if (_styles.TryGetValue(styleId, out s))
                return s;
            return _byName.TryGetValue(styleId, out s) ? s : null;
        }

        private IEnumerable<WordStyle> Chain(string styleId)
        {
            var chain = new List<WordStyle>();
            var guard = 0;
            var current = Get(styleId);
            while (current != null && guard++ < 32)
            {
                chain.Add(current);
                current = Get(current.BasedOn);
            }
            chain.Reverse();     // root first so derived styles win
            return chain;
        }

        /// <summary>Paragraph formatting contributed by a style chain (document defaults not included).</summary>
        public ParagraphFormat ParagraphFormatFor(string styleId)
        {
            if (string.IsNullOrEmpty(styleId))
                return new ParagraphFormat();
            ParagraphFormat cached;
            if (_pCache.TryGetValue(styleId, out cached))
                return cached;

            var result = new ParagraphFormat();
            foreach (var style in Chain(styleId))
            {
                if (style.Type == "paragraph" || style.Type == "table")
                    result.ApplyOver(_fmt.ReadParagraphFormat(style.PPr));
            }
            _pCache[styleId] = result;
            return result;
        }

        /// <summary>Character formatting contributed by a style chain (document defaults not included).</summary>
        public CharacterFormat CharacterFormatFor(string styleId)
        {
            if (string.IsNullOrEmpty(styleId))
                return new CharacterFormat();
            CharacterFormat cached;
            if (_rCache.TryGetValue(styleId, out cached))
                return cached;

            var result = new CharacterFormat();
            foreach (var style in Chain(styleId))
                result.ApplyOver(_fmt.ReadRunFormat(style.RPr));
            _rCache[styleId] = result;
            return result;
        }

        /// <summary>
        /// Numbering reference declared by a paragraph style or an ancestor of it.
        /// A derived style often declares only the level (w:ilvl) and inherits the list itself,
        /// so the nearest declared level wins.
        /// </summary>
        public bool TryGetStyleNumbering(string styleId, out string numId, out int level)
        {
            numId = null;
            level = 0;
            int? nearestLevel = null;

            foreach (var style in Chain(styleId).Reverse())     // leaf first
            {
                if (!nearestLevel.HasValue && style.NumLevel.HasValue)
                    nearestLevel = style.NumLevel;
                if (!string.IsNullOrEmpty(style.NumId))
                {
                    numId = style.NumId;
                    level = nearestLevel ?? style.NumLevel ?? 0;
                    return true;
                }
            }
            return false;
        }

        public int HeadingLevelOf(string styleId)
        {
            var style = Get(styleId);
            if (style == null)
                return 0;

            foreach (var s in Chain(styleId).Reverse())
            {
                var name = s.Name ?? s.Id;
                if (string.IsNullOrEmpty(name))
                    continue;
                var compact = name.Replace(" ", string.Empty);
                if (compact.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
                {
                    int lvl;
                    if (int.TryParse(compact.Substring(7), out lvl) && lvl >= 1 && lvl <= 9)
                        return lvl;
                }
                if (s.OutlineLevel.HasValue && s.OutlineLevel.Value >= 0 && s.OutlineLevel.Value <= 8)
                    return s.OutlineLevel.Value + 1;
            }
            return 0;
        }

        /// <summary>Table-level properties (borders, cell margins) contributed by a table style chain.</summary>
        public IEnumerable<XElement> TableStyleProperties(string styleId)
        {
            foreach (var style in Chain(styleId))
            {
                if (style.TblPr != null)
                    yield return style.TblPr;
            }
        }

        public IEnumerable<XElement> TableStyleCellProperties(string styleId)
        {
            foreach (var style in Chain(styleId))
            {
                if (style.TcPr != null)
                    yield return style.TcPr;
            }
        }

        /// <summary>
        /// The w:tblStylePr elements for one conditional region (firstRow, band1Horz, ...)
        /// across the style chain, root first so derived styles win when applied in order.
        /// </summary>
        public IEnumerable<XElement> TableStyleRegions(string styleId, string region)
        {
            foreach (var style in Chain(styleId))
            {
                if (style.TblStylePrs == null)
                    continue;
                foreach (var part in style.TblStylePrs)
                {
                    if (string.Equals(OoxmlUtil.Str(part, Ns.W + "type"), region, StringComparison.OrdinalIgnoreCase))
                        yield return part;
                }
            }
        }
    }
}
