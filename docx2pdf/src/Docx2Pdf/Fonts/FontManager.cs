using System;
using System.Collections.Generic;

namespace Docx2Pdf.Fonts
{
    /// <summary>A stretch of text that can be drawn with a single font.</summary>
    internal struct FontRun
    {
        public PdfFontBase Font;
        public string Text;
    }

    /// <summary>
    /// Maps document font requests to PDF fonts: embeds the real font file when it can be found,
    /// substitutes a base-14 font otherwise, and picks fallback fonts for unsupported characters.
    /// </summary>
    internal sealed class FontManager
    {
        private readonly ConversionOptions _options;
        private readonly List<string> _warnings;
        private readonly Dictionary<string, PdfFontBase> _byRequest = new Dictionary<string, PdfFontBase>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, PdfFontBase> _byFile = new Dictionary<string, PdfFontBase>(StringComparer.OrdinalIgnoreCase);
        private readonly List<PdfFontBase> _fonts = new List<PdfFontBase>();
        private readonly List<PdfFontBase> _fallbackCache = new List<PdfFontBase>();
        private bool _fallbacksResolvedFor;
        private bool _fallbackBold, _fallbackItalic;

        public FontManager(ConversionOptions options, List<string> warnings)
        {
            _options = options ?? new ConversionOptions();
            _warnings = warnings ?? new List<string>();

            if (_options.EmbedFonts)
            {
                var dirs = new List<string>(_options.FontDirectories);
                dirs.AddRange(SystemFontIndex.DefaultDirectories());
                try
                {
                    SystemFontIndex.Build(dirs);
                }
                catch (Exception ex)
                {
                    Warn("Font directories could not be scanned (" + ex.Message + "); standard fonts will be used.");
                }
            }
        }

        public IList<PdfFontBase> Fonts { get { return _fonts; } }

        private void Warn(string message)
        {
            if (!_warnings.Contains(message))
                _warnings.Add(message);
        }

        public PdfFontBase Resolve(string family, bool bold, bool italic)
        {
            if (string.IsNullOrEmpty(family))
                family = _options.DefaultFontFamily ?? "Calibri";

            var key = family + "|" + (bold ? "b" : string.Empty) + (italic ? "i" : string.Empty);
            PdfFontBase font;
            if (_byRequest.TryGetValue(key, out font))
                return font;

            font = LoadEmbedded(family, bold, italic);

            if (font == null && !string.Equals(family, _options.DefaultFontFamily, StringComparison.OrdinalIgnoreCase))
            {
                font = LoadEmbedded(_options.DefaultFontFamily, bold, italic);
                if (font != null)
                    Warn("Font '" + family + "' is not installed; '" + _options.DefaultFontFamily + "' was used instead.");
            }

            if (font == null)
            {
                var standard = StandardFontMetrics.ClassifyFamily(family);
                font = Register(new StandardPdfFont(standard, bold, italic));
                if (_options.EmbedFonts)
                    Warn("Font '" + family + "' is not installed; the standard PDF font '" + font.DisplayName + "' was substituted.");
            }

            _byRequest[key] = font;
            return font;
        }

        private PdfFontBase LoadEmbedded(string family, bool bold, bool italic)
        {
            if (!_options.EmbedFonts || string.IsNullOrEmpty(family))
                return null;

            var entry = SystemFontIndex.Find(family, bold, italic);
            if (entry == null)
                return null;

            var fileKey = entry.Path + "#" + entry.Index;
            PdfFontBase cached;
            if (_byFile.TryGetValue(fileKey, out cached))
                return cached;

            try
            {
                var file = TrueTypeFile.Load(entry.Path, entry.Index);
                if (!file.HasGlyphOutlines)
                    return null;
                var font = Register(new EmbeddedPdfFont(file));
                _byFile[fileKey] = font;
                return font;
            }
            catch (Exception ex)
            {
                Warn("Font file '" + entry.Path + "' could not be embedded (" + ex.Message + ").");
                return null;
            }
        }

        private PdfFontBase Register(PdfFontBase font)
        {
            _fonts.Add(font);
            font.ResourceName = "F" + _fonts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return font;
        }

        /// <summary>
        /// Splits text into runs that share a single font, switching to a fallback font
        /// for characters the primary font cannot render.
        /// </summary>
        public List<FontRun> Split(string text, PdfFontBase primary, bool bold, bool italic)
        {
            var runs = new List<FontRun>();
            if (string.IsNullOrEmpty(text))
                return runs;

            PdfFontBase current = null;
            var start = 0;
            var i = 0;
            while (i < text.Length)
            {
                var length = 1;
                int cp = text[i];
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    cp = char.ConvertToUtf32(text[i], text[i + 1]);
                    length = 2;
                }

                var font = FontFor(cp, primary, bold, italic);
                if (current == null)
                {
                    current = font;
                    start = i;
                }
                else if (!ReferenceEquals(font, current))
                {
                    runs.Add(new FontRun { Font = current, Text = text.Substring(start, i - start) });
                    current = font;
                    start = i;
                }
                i += length;
            }

            if (current != null && start < text.Length)
                runs.Add(new FontRun { Font = current, Text = text.Substring(start) });
            return runs;
        }

        private PdfFontBase FontFor(int codePoint, PdfFontBase primary, bool bold, bool italic)
        {
            if (codePoint == ' ' || primary.Supports(codePoint))
                return primary;

            foreach (var fallback in Fallbacks(bold, italic))
            {
                if (fallback.Supports(codePoint))
                    return fallback;
            }
            return primary;
        }

        private IEnumerable<PdfFontBase> Fallbacks(bool bold, bool italic)
        {
            if (!_fallbacksResolvedFor || _fallbackBold != bold || _fallbackItalic != italic)
            {
                _fallbackCache.Clear();
                foreach (var family in _options.FallbackFontFamilies)
                {
                    var font = LoadEmbedded(family, bold, italic);
                    if (font != null)
                        _fallbackCache.Add(font);
                }
                _fallbacksResolvedFor = true;
                _fallbackBold = bold;
                _fallbackItalic = italic;
            }
            return _fallbackCache;
        }

        /// <summary>Names of the fonts that will be embedded in the output.</summary>
        public IList<string> EmbeddedFontNames()
        {
            var names = new List<string>();
            foreach (var font in _fonts)
            {
                if (font is EmbeddedPdfFont && !names.Contains(font.DisplayName))
                    names.Add(font.DisplayName);
            }
            return names;
        }
    }
}
