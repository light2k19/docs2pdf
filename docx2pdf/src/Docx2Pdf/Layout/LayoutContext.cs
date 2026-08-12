using System;
using System.Collections.Generic;
using Docx2Pdf.Fonts;
using Docx2Pdf.Images;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    /// <summary>Shared services and caches used while laying a document out.</summary>
    internal sealed class LayoutContext
    {
        public readonly FontManager Fonts;
        public readonly ConversionOptions Options;
        public readonly List<string> Warnings;
        public double DefaultTabStopPt = 36;
        /// <summary>Legacy compatibility mode (&lt; 15): cell paragraphs that do not follow
        /// another paragraph drop their space-before (sample1's nested-table demo).</summary>
        public bool LegacyCellSpacing;
        /// <summary>Style id of the document's default paragraph style ("Normal").</summary>
        public string DefaultParagraphStyleId;

        private readonly Dictionary<string, DecodedImage> _images = new Dictionary<string, DecodedImage>(StringComparer.Ordinal);
        private readonly HashSet<string> _failedImages = new HashSet<string>(StringComparer.Ordinal);

        public LayoutContext(FontManager fonts, ConversionOptions options, List<string> warnings)
        {
            Fonts = fonts;
            Options = options;
            Warnings = warnings;
        }

        public void Warn(string message)
        {
            if (!Warnings.Contains(message))
                Warnings.Add(message);
        }

        /// <summary>Decodes an image once per package part.</summary>
        public DecodedImage GetImage(ImageInline image)
        {
            if (image == null || image.Data == null)
                return null;
            var key = image.PartName ?? ("anon:" + image.Data.Length);
            DecodedImage decoded;
            if (_images.TryGetValue(key, out decoded))
                return decoded;
            if (_failedImages.Contains(key))
                return null;

            string format;
            try
            {
                decoded = ImageDecoder.Decode(image.Data, out format);
            }
            catch (Exception ex)
            {
                decoded = null;
                format = "image";
                Warn("An image could not be decoded (" + ex.Message + ").");
            }

            if (decoded == null)
            {
                _failedImages.Add(key);
                Warn(format + " images are not supported and were left blank.");
                return null;
            }

            _images[key] = decoded;
            return decoded;
        }

        public IEnumerable<KeyValuePair<string, DecodedImage>> Images { get { return _images; } }

        public PdfFontBase ResolveFont(CharacterFormat format)
        {
            var family = format != null && !string.IsNullOrEmpty(format.FontFamily)
                ? format.FontFamily
                : Options.DefaultFontFamily;
            var bold = format != null && format.Bold == true;
            var italic = format != null && format.Italic == true;
            return Fonts.Resolve(family, bold, italic);
        }
    }
}
