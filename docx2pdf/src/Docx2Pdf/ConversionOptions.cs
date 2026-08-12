using System.Collections.Generic;

namespace Docx2Pdf
{
    /// <summary>Settings that control how a DOCX file is rendered to PDF.</summary>
    public sealed class ConversionOptions
    {
        /// <summary>
        /// Embed the document's own fonts by locating the matching font files on the machine.
        /// When false (or when a font cannot be found) the standard PDF fonts are substituted.
        /// </summary>
        public bool EmbedFonts { get; set; }

        /// <summary>Extra directories to search for font files, in addition to the platform defaults.</summary>
        public IList<string> FontDirectories { get; private set; }

        /// <summary>
        /// Fonts tried, in order, for characters the requested font cannot render
        /// (for example CJK text in a document that asks for Calibri).
        /// </summary>
        public IList<string> FallbackFontFamilies { get; private set; }

        /// <summary>Font used when a document does not name one.</summary>
        public string DefaultFontFamily { get; set; }

        /// <summary>Compress page content and image streams with FlateDecode.</summary>
        public bool CompressStreams { get; set; }

        /// <summary>Emit a PDF outline (bookmarks) built from the document's heading styles.</summary>
        public bool GenerateOutline { get; set; }

        /// <summary>Emit link annotations for hyperlinks.</summary>
        public bool CreateHyperlinks { get; set; }

        /// <summary>Render page headers and footers.</summary>
        public bool RenderHeadersAndFooters { get; set; }

        /// <summary>Render images. Turning this off produces a text-only PDF.</summary>
        public bool RenderImages { get; set; }

        /// <summary>Overrides the document title stored in the PDF metadata.</summary>
        public string Title { get; set; }

        /// <summary>Overrides the author stored in the PDF metadata.</summary>
        public string Author { get; set; }

        /// <summary>Value written to the PDF /Producer entry.</summary>
        public string Producer { get; set; }

        /// <summary>Hard limit on generated pages; guards against pathological documents.</summary>
        public int MaxPages { get; set; }

        public ConversionOptions()
        {
            EmbedFonts = true;
            CompressStreams = true;
            GenerateOutline = true;
            CreateHyperlinks = true;
            RenderHeadersAndFooters = true;
            RenderImages = true;
            DefaultFontFamily = "Calibri";
            Producer = "Docx2Pdf";
            MaxPages = 20000;
            FontDirectories = new List<string>();
            FallbackFontFamilies = new List<string>
            {
                "Arial",
                "Segoe UI",
                "Times New Roman",
                "Segoe UI Symbol",
                "Microsoft YaHei",
                "SimSun",
                "MS Gothic",
                "Malgun Gothic",
                "Nirmala UI",
                "DejaVu Sans",
                "Noto Sans",
                "FreeSerif",
            };
        }
    }
}
