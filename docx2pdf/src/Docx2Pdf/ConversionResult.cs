using System.Collections.Generic;

namespace Docx2Pdf
{
    /// <summary>Outcome of a conversion: what was produced and anything that could not be rendered exactly.</summary>
    public sealed class ConversionResult
    {
        /// <summary>Number of pages written.</summary>
        public int PageCount { get; internal set; }

        /// <summary>Number of bytes written to the output stream.</summary>
        public long ByteCount { get; internal set; }

        /// <summary>Fonts embedded into the PDF.</summary>
        public IList<string> EmbeddedFonts { get; internal set; }

        /// <summary>Non-fatal fidelity notes (unsupported features, substituted fonts, skipped images).</summary>
        public IList<string> Warnings { get; internal set; }

        internal ConversionResult()
        {
            EmbeddedFonts = new List<string>();
            Warnings = new List<string>();
        }
    }
}
