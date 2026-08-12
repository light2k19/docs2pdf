using System;
using System.Collections.Generic;
using System.IO;
using Docx2Pdf.Fonts;
using Docx2Pdf.Layout;
using Docx2Pdf.Ooxml;
using Docx2Pdf.Pdf;

namespace Docx2Pdf
{
    /// <summary>
    /// Converts WordprocessingML (.docx) documents to PDF.
    /// The implementation is pure managed code with no external dependencies and no Microsoft Word requirement.
    /// </summary>
    public static class DocxToPdfConverter
    {
        /// <summary>Converts a .docx file to a .pdf file.</summary>
        /// <param name="docxPath">Path of the source .docx file.</param>
        /// <param name="pdfPath">Path of the .pdf file to create (overwritten when it exists).</param>
        /// <param name="options">Optional conversion settings.</param>
        public static ConversionResult Convert(string docxPath, string pdfPath, ConversionOptions options = null)
        {
            if (string.IsNullOrEmpty(docxPath))
                throw new ArgumentNullException("docxPath");
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException("pdfPath");
            if (!File.Exists(docxPath))
                throw new FileNotFoundException("The source document was not found.", docxPath);

            var directory = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using (var input = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var output = new FileStream(pdfPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                return Convert(input, output, options);
            }
        }

        /// <summary>Converts a .docx stream to PDF, writing the result to <paramref name="pdfStream"/>.</summary>
        public static ConversionResult Convert(Stream docxStream, Stream pdfStream, ConversionOptions options = null)
        {
            if (docxStream == null)
                throw new ArgumentNullException("docxStream");
            if (pdfStream == null)
                throw new ArgumentNullException("pdfStream");

            options = options ?? new ConversionOptions();
            var warnings = new List<string>();
            var result = new ConversionResult();

            // ZipArchive needs a seekable stream.
            var source = docxStream;
            MemoryStream buffered = null;
            if (!docxStream.CanSeek)
            {
                buffered = new MemoryStream();
                docxStream.CopyTo(buffered);
                buffered.Position = 0;
                source = buffered;
            }

            try
            {
                using (var package = OpcPackage.Open(source, true))
                {
                    var reader = new DocumentReader(package, warnings);
                    var document = reader.Read();

                    var fonts = new FontManager(options, warnings);
                    var context = new LayoutContext(fonts, options, warnings);
                    var layout = new DocumentLayout(context).Layout(document);

                    var renderer = new PdfRenderer(options, fonts);
                    var pdf = renderer.Render(layout, document.Info);

                    var counter = new CountingStream(pdfStream);
                    pdf.Save(counter);
                    counter.Flush();

                    result.PageCount = layout.Pages.Count;
                    result.ByteCount = counter.BytesWritten;
                    result.EmbeddedFonts = fonts.EmbeddedFontNames();
                }
            }
            finally
            {
                if (buffered != null)
                    buffered.Dispose();
            }

            result.Warnings = warnings;
            return result;
        }

        /// <summary>Converts .docx bytes to PDF bytes.</summary>
        public static byte[] Convert(byte[] docxBytes, ConversionOptions options = null)
        {
            if (docxBytes == null)
                throw new ArgumentNullException("docxBytes");
            using (var input = new MemoryStream(docxBytes, false))
            using (var output = new MemoryStream())
            {
                Convert(input, output, options);
                return output.ToArray();
            }
        }

        private sealed class CountingStream : Stream
        {
            private readonly Stream _inner;
            public long BytesWritten;

            public CountingStream(Stream inner) { _inner = inner; }

            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { return BytesWritten; } }
            public override long Position
            {
                get { return BytesWritten; }
                set { throw new NotSupportedException(); }
            }

            public override void Flush() { _inner.Flush(); }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }

            public override void Write(byte[] buffer, int offset, int count)
            {
                _inner.Write(buffer, offset, count);
                BytesWritten += count;
            }
        }
    }
}
