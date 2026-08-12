using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Docx2Pdf.Pdf;

namespace Docx2Pdf.Fonts
{
    /// <summary>A font as used by the layout engine and written into the PDF resource dictionary.</summary>
    internal abstract class PdfFontBase
    {
        /// <summary>Resource name inside the page resource dictionary (F1, F2, ...).</summary>
        public string ResourceName;
        public string DisplayName;

        public abstract double AscentEm { get; }
        public abstract double DescentEm { get; }     // negative
        public abstract double LineHeightEm { get; }
        public abstract bool Supports(int codePoint);
        public abstract double WidthEm(int codePoint);
        public abstract void AppendEncoded(StringBuilder hex, string text);

        public abstract PdfObject Build(PdfDocument doc);

        public double Measure(string text, double sizePt)
        {
            if (string.IsNullOrEmpty(text))
                return 0;
            double total = 0;
            for (var i = 0; i < text.Length; i++)
            {
                int cp = text[i];
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    cp = char.ConvertToUtf32(text[i], text[i + 1]);
                    i++;
                }
                total += WidthEm(cp);
            }
            return total * sizePt;
        }

        public string EncodeHex(string text)
        {
            var sb = new StringBuilder();
            AppendEncoded(sb, text);
            return sb.ToString();
        }

        protected static IEnumerable<int> CodePoints(string text)
        {
            for (var i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    yield return char.ConvertToUtf32(text[i], text[i + 1]);
                    i++;
                }
                else
                {
                    yield return text[i];
                }
            }
        }
    }

    /// <summary>One of the 14 standard PDF fonts, used when no font file can be embedded.</summary>
    internal sealed class StandardPdfFont : PdfFontBase
    {
        private readonly short[] _widths;
        private readonly string _baseFont;

        public StandardPdfFont(StandardFontFamily family, bool bold, bool italic)
        {
            _widths = StandardFontMetrics.GetWidths(family, bold, italic);
            _baseFont = StandardFontMetrics.BaseFontName(family, bold, italic);
            DisplayName = _baseFont;
        }

        public override double AscentEm { get { return 0.75; } }
        public override double DescentEm { get { return -0.25; } }
        public override double LineHeightEm { get { return 1.15; } }

        public override bool Supports(int codePoint)
        {
            byte b;
            return WinAnsiEncoding.TryGetByte(codePoint, out b);
        }

        public override double WidthEm(int codePoint)
        {
            byte b;
            if (!WinAnsiEncoding.TryGetByte(codePoint, out b))
                b = (byte)'?';
            return _widths[b] / 1000.0;
        }

        public override void AppendEncoded(StringBuilder hex, string text)
        {
            foreach (var cp in CodePoints(text))
            {
                byte b;
                if (!WinAnsiEncoding.TryGetByte(cp, out b))
                    b = (byte)'?';
                hex.Append(b.ToString("X2", CultureInfo.InvariantCulture));
            }
        }

        public override PdfObject Build(PdfDocument doc)
        {
            var dict = new PdfDictionary();
            dict.Set("Type", "Font");
            dict.Set("Subtype", "Type1");
            dict.Set("BaseFont", _baseFont);
            dict.Set("Encoding", "WinAnsiEncoding");
            return doc.Add(dict);
        }
    }

    /// <summary>An embedded TrueType/OpenType font written as a Type0 (Identity-H) composite font.</summary>
    internal sealed class EmbeddedPdfFont : PdfFontBase
    {
        private readonly TrueTypeFile _font;
        private readonly Dictionary<int, int> _glyphToUnicode = new Dictionary<int, int>();
        private readonly Dictionary<int, int> _glyphCache = new Dictionary<int, int>();
        private readonly double _scale;

        public EmbeddedPdfFont(TrueTypeFile font)
        {
            _font = font;
            _scale = 1.0 / font.UnitsPerEm;
            DisplayName = font.PostScriptName ?? font.FamilyName ?? "Embedded";
        }

        public TrueTypeFile File { get { return _font; } }

        public override double AscentEm
        {
            get
            {
                // A font that sets OS/2 USE_TYPO_METRICS asks for the sTypo values (Lato:
                // win metrics are 18% taller and blew a 2-page resume up to 3 pages).
                if (_font.UseTypoMetrics && _font.TypoAscender != 0)
                    return _font.TypoAscender * _scale;
                // Otherwise Word (via GDI) places the baseline using the OS/2 usWinAscent,
                // not the hhea ascender — for Calibri the two differ by a fifth of an em.
                if (_font.WinAscent > 0)
                    return _font.WinAscent * _scale;
                var ascent = _font.Ascender != 0 ? _font.Ascender : _font.TypoAscender;
                if (ascent == 0)
                    ascent = (short)(_font.UnitsPerEm * 0.8);
                return ascent * _scale;
            }
        }

        public override double DescentEm
        {
            get
            {
                if (_font.UseTypoMetrics && _font.TypoDescender != 0)
                    return _font.TypoDescender * _scale;
                if (_font.WinDescent > 0)
                    return -(_font.WinDescent * _scale);
                var descent = _font.Descender != 0 ? _font.Descender : _font.TypoDescender;
                if (descent == 0)
                    descent = (short)(-_font.UnitsPerEm * 0.2);
                return descent * _scale;
            }
        }

        public override double LineHeightEm
        {
            get
            {
                if (_font.UseTypoMetrics && _font.TypoAscender != 0)
                {
                    var typo = (_font.TypoAscender - _font.TypoDescender + Math.Max((int)_font.TypoLineGap, 0)) * _scale;
                    return typo < 1.0 ? 1.0 : typo;
                }
                // GDI line height: usWinAscent + usWinDescent plus external leading, where
                // the leading is the amount by which the hhea metrics exceed the win metrics.
                var winHeight = AscentEm - DescentEm;
                var hheaHeight = (Math.Abs(_font.Ascender) + Math.Abs(_font.Descender) + Math.Max((int)_font.LineGap, 0)) * _scale;
                var height = winHeight + Math.Max(0, hheaHeight - winHeight);
                return height < 1.0 ? 1.0 : height;
            }
        }

        private int GlyphOf(int codePoint)
        {
            int gid;
            if (_glyphCache.TryGetValue(codePoint, out gid))
                return gid;
            gid = _font.GetGlyphId(codePoint);
            _glyphCache[codePoint] = gid;
            return gid;
        }

        public override bool Supports(int codePoint)
        {
            if (codePoint == '\t' || codePoint == '\n' || codePoint == '\r')
                return true;
            return GlyphOf(codePoint) != 0;
        }

        public override double WidthEm(int codePoint)
        {
            var gid = GlyphOf(codePoint);
            if (gid == 0 && codePoint == ' ')
                return 0.25;
            return _font.GetAdvance(gid) * _scale;
        }

        public override void AppendEncoded(StringBuilder hex, string text)
        {
            foreach (var cp in CodePoints(text))
            {
                var gid = GlyphOf(cp);
                if (gid == 0)
                    gid = GlyphOf('?');
                if (!_glyphToUnicode.ContainsKey(gid))
                    _glyphToUnicode[gid] = cp;
                hex.Append(gid.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        public override PdfObject Build(PdfDocument doc)
        {
            var baseFont = SanitizeName(_font.PostScriptName ?? _font.FamilyName ?? "Font");

            // Embed a subset with only the glyphs the document uses (ids stay stable, so
            // the Identity CID mapping is unaffected). A full Arial program is ~1MB; the
            // subset is a few KB. Subset names carry the conventional 6-letter tag.
            var fontData = _font.Data;
            if (!_font.IsCff)
            {
                try
                {
                    fontData = _font.BuildSubset(_glyphToUnicode.Keys);
                    if (!ReferenceEquals(fontData, _font.Data))
                        baseFont = SubsetTag(baseFont) + "+" + baseFont;
                }
                catch
                {
                    fontData = _font.Data;
                }
            }

            var descriptor = new PdfDictionary();
            descriptor.Set("Type", "FontDescriptor");
            descriptor.Set("FontName", baseFont);
            descriptor.Set("Flags", BuildFlags());
            var scale = 1000.0 / _font.UnitsPerEm;
            descriptor.Set("FontBBox", new PdfArray()
                .Add(_font.XMin * scale).Add(_font.YMin * scale)
                .Add(_font.XMax * scale).Add(_font.YMax * scale));
            descriptor.Set("ItalicAngle", _font.ItalicAngle);
            descriptor.Set("Ascent", AscentEm * 1000);
            descriptor.Set("Descent", DescentEm * 1000);
            descriptor.Set("CapHeight", _font.CapHeight != 0 ? _font.CapHeight * scale : AscentEm * 1000);
            descriptor.Set("StemV", _font.IsBold ? 160 : 80);

            var program = new PdfStream(Flate.Compress(fontData));
            program.Set("Filter", "FlateDecode");
            if (_font.IsCff)
                program.Set("Subtype", "OpenType");
            else
                program.Set("Length1", fontData.Length);
            descriptor.Set(_font.IsCff ? "FontFile3" : "FontFile2", doc.Add(program));

            var descendant = new PdfDictionary();
            descendant.Set("Type", "Font");
            descendant.Set("Subtype", _font.IsCff ? "CIDFontType0" : "CIDFontType2");
            descendant.Set("BaseFont", baseFont);
            var systemInfo = new PdfDictionary();
            systemInfo.Set("Registry", new PdfString("Adobe"));
            systemInfo.Set("Ordering", new PdfString("Identity"));
            systemInfo.Set("Supplement", 0);
            descendant.Set("CIDSystemInfo", systemInfo);
            descendant.Set("FontDescriptor", doc.Add(descriptor));
            descendant.Set("DW", 1000);
            descendant.Set("W", BuildWidths());
            if (!_font.IsCff)
                descendant.Set("CIDToGIDMap", "Identity");

            var type0 = new PdfDictionary();
            type0.Set("Type", "Font");
            type0.Set("Subtype", "Type0");
            type0.Set("BaseFont", baseFont);
            type0.Set("Encoding", "Identity-H");
            type0.Set("DescendantFonts", new PdfArray(doc.Add(descendant)));
            type0.Set("ToUnicode", doc.Add(BuildToUnicode()));
            return doc.Add(type0);
        }

        private int BuildFlags()
        {
            var flags = 0;
            if (_font.IsFixedPitch) flags |= 1;
            if (_font.IsSerif) flags |= 2;
            flags |= 32;                       // nonsymbolic
            if (_font.IsItalic || Math.Abs(_font.ItalicAngle) > 0.01) flags |= 64;
            return flags;
        }

        private PdfArray BuildWidths()
        {
            var gids = new List<int>(_glyphToUnicode.Keys);
            gids.Sort();
            var array = new PdfArray();
            var scale = 1000.0 / _font.UnitsPerEm;

            var i = 0;
            while (i < gids.Count)
            {
                var start = gids[i];
                var group = new PdfArray();
                var previous = start - 1;
                while (i < gids.Count && gids[i] == previous + 1)
                {
                    group.Add(Math.Round(_font.GetAdvance(gids[i]) * scale, 1));
                    previous = gids[i];
                    i++;
                }
                array.Add(start);
                array.Add(group);
            }
            return array;
        }

        private PdfStream BuildToUnicode()
        {
            var sb = new StringBuilder();
            sb.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
            sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            sb.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
            sb.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");

            var entries = new List<KeyValuePair<int, int>>(_glyphToUnicode);
            entries.Sort((a, b) => a.Key.CompareTo(b.Key));
            for (var offset = 0; offset < entries.Count; offset += 100)
            {
                var count = Math.Min(100, entries.Count - offset);
                sb.Append(count.ToString(CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
                for (var i = offset; i < offset + count; i++)
                {
                    sb.Append('<').Append(entries[i].Key.ToString("X4", CultureInfo.InvariantCulture)).Append("> <");
                    var cp = entries[i].Value;
                    if (cp > 0xFFFF)
                    {
                        var s = char.ConvertFromUtf32(cp);
                        foreach (var c in s)
                            sb.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        sb.Append(cp.ToString("X4", CultureInfo.InvariantCulture));
                    }
                    sb.Append(">\n");
                }
                sb.Append("endbfchar\n");
            }
            sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");

            var stream = new PdfStream(Flate.Compress(Encoding.ASCII.GetBytes(sb.ToString())));
            stream.Set("Filter", "FlateDecode");
            return stream;
        }

        private static string SanitizeName(string name)
        {
            var sb = new StringBuilder();
            foreach (var c in name)
            {
                if (char.IsLetterOrDigit(c) || c == '-' || c == '+' || c == '.' || c == '_')
                    sb.Append(c);
            }
            return sb.Length == 0 ? "Font" : sb.ToString();
        }

        /// <summary>Deterministic six-letter subset tag derived from the font name.</summary>
        private static string SubsetTag(string name)
        {
            var hash = 5381u;
            foreach (var c in name)
                hash = unchecked(hash * 33 + c);
            var tag = new char[6];
            for (var i = 0; i < 6; i++)
            {
                tag[i] = (char)('A' + (int)(hash % 26));
                hash /= 26;
            }
            return new string(tag);
        }
    }
}
