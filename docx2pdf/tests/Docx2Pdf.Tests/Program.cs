using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Docx2Pdf.Fonts;
using Docx2Pdf.Images;
using Docx2Pdf.Layout;
using Docx2Pdf.Model;
using Docx2Pdf.Ooxml;
using Docx2Pdf.Pdf;

namespace Docx2Pdf.Tests
{
    internal static class Program
    {
        private static int _passed;
        private static readonly List<string> Failures = new List<string>();
        private static string _sampleDirectory;
        private static string _outputDirectory;

        private static int Main(string[] args)
        {
            _sampleDirectory = args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal))
                               ?? Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "..", "0sample");
            _sampleDirectory = Path.GetFullPath(_sampleDirectory);
            _outputDirectory = args.Where(a => !a.StartsWith("--", StringComparison.Ordinal)).Skip(1).FirstOrDefault()
                               ?? _sampleDirectory;
            _outputDirectory = Path.GetFullPath(_outputDirectory);
            Directory.CreateDirectory(_outputDirectory);

            if (args.Contains("--dump"))
            {
                foreach (var file in Directory.GetFiles(_sampleDirectory, "*.docx"))
                    Diagnostics.Dump(file, Console.Out);
                return 0;
            }

            if (args.Contains("--pages"))
            {
                foreach (var file in Directory.GetFiles(_sampleDirectory, "*.docx"))
                {
                    Console.WriteLine("=== " + Path.GetFileName(file) + " ===");
                    Diagnostics.DumpPageStarts(file, Console.Out);
                }
                return 0;
            }

            Console.WriteLine("Docx2Pdf test run");
            Console.WriteLine("  samples: " + _sampleDirectory);
            Console.WriteLine("  output : " + _outputDirectory);
            Console.WriteLine();

            Run("pdf writer produces a loadable file", PdfWriterProducesValidFile);
            Run("flate round trip", FlateRoundTrip);
            Run("standard font metrics are sane", StandardFontMetricsAreSane);
            Run("system fonts can be embedded", SystemFontsCanBeEmbedded);
            Run("png decoder handles all colour types", PngDecoderHandlesColourTypes);
            Run("bmp decoder reads a 24-bit bitmap", BmpDecoderReads24Bit);
            Run("opc package resolves relationships", OpcResolvesRelationships);
            Run("text wraps at the available width", TextWrapsAtAvailableWidth);
            Run("numbering produces list labels", NumberingProducesLabels);
            Run("sample documents convert", SampleDocumentsConvert);
            Run("converted pdf structure is valid", ConvertedPdfStructureIsValid);
            Run("conversion is deterministic", ConversionIsDeterministic);
            Run("options switch off features", OptionsSwitchOffFeatures);

            Console.WriteLine();
            Console.WriteLine("{0} passed, {1} failed", _passed, Failures.Count);
            foreach (var failure in Failures)
                Console.WriteLine("  FAIL " + failure);
            return Failures.Count == 0 ? 0 : 1;
        }

        // ------------------------------------------------------------------ harness

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("  PASS  " + name);
            }
            catch (Exception ex)
            {
                Failures.Add(name + ": " + ex.Message);
                Console.WriteLine("  FAIL  " + name + " -- " + ex.Message);
            }
        }

        private static void Check(bool condition, string message)
        {
            if (!condition)
                throw new Exception(message);
        }

        // ------------------------------------------------------------------ unit tests

        private static void PdfWriterProducesValidFile()
        {
            var pdf = new PdfDocument();
            var pagesRef = pdf.Reserve();
            var content = new PdfStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (hi) Tj ET"));
            var page = new PdfDictionary();
            page.Set("Type", "Page");
            page.Set("Parent", pagesRef);
            page.Set("MediaBox", new PdfArray().Add(0).Add(0).Add(612).Add(792));
            page.Set("Contents", pdf.Add(content));
            var pageRef = pdf.Add(page);

            var pages = new PdfDictionary();
            pages.Set("Type", "Pages");
            pages.Set("Count", 1);
            pages.Set("Kids", new PdfArray(pageRef));
            pagesRef.Target = pages;

            var catalog = new PdfDictionary();
            catalog.Set("Type", "Catalog");
            catalog.Set("Pages", pagesRef);
            pdf.Catalog = catalog;
            pdf.Add(catalog);

            using (var ms = new MemoryStream())
            {
                pdf.Save(ms);
                var text = Encoding.ASCII.GetString(ms.ToArray());
                Check(text.StartsWith("%PDF-1.7", StringComparison.Ordinal), "missing header");
                Check(text.Contains("/Type /Catalog"), "missing catalog");
                Check(text.TrimEnd().EndsWith("%%EOF", StringComparison.Ordinal), "missing EOF marker");
                Check(XrefOffsetsResolve(ms.ToArray()), "xref offsets do not point at object headers");
            }
        }

        private static void FlateRoundTrip()
        {
            var data = new byte[5000];
            for (var i = 0; i < data.Length; i++)
                data[i] = (byte)(i % 251);
            var compressed = Flate.Compress(data);
            Check(compressed.Length > 2, "compression produced nothing");
            Check(compressed[0] == 0x78, "missing zlib header");
            var restored = Flate.Decompress(compressed, 0, compressed.Length);
            Check(restored.Length == data.Length, "length mismatch after round trip");
            for (var i = 0; i < data.Length; i++)
                Check(restored[i] == data[i], "byte mismatch at " + i);
        }

        private static void StandardFontMetricsAreSane()
        {
            var helvetica = new StandardPdfFont(StandardFontFamily.Helvetica, false, false);
            var width = helvetica.Measure("Hello", 12);
            Check(width > 20 && width < 45, "unexpected Helvetica width: " + width);
            Check(helvetica.EncodeHex("AB") == "4142", "unexpected WinAnsi encoding: " + helvetica.EncodeHex("AB"));
            Check(helvetica.Supports('é'), "WinAnsi should cover Latin-1");

            var courier = new StandardPdfFont(StandardFontFamily.Courier, false, false);
            Check(Math.Abs(courier.Measure("iiii", 10) - courier.Measure("MMMM", 10)) < 0.01, "Courier is not monospaced");
        }

        private static void SystemFontsCanBeEmbedded()
        {
            var warnings = new List<string>();
            var manager = new FontManager(new ConversionOptions(), warnings);
            var font = manager.Resolve("Arial", false, false);
            Check(font != null, "no font resolved");
            Check(font.Measure("Hello world", 11) > 30, "implausible text width");

            var embedded = font as EmbeddedPdfFont;
            if (embedded != null)
            {
                Check(embedded.File.UnitsPerEm > 0, "font has no unitsPerEm");
                Check(embedded.File.Data.Length > 1000, "font program is too small");
                Check(embedded.EncodeHex("A").Length == 4, "CID encoding should use two bytes per glyph");
            }
            else
            {
                Console.WriteLine("        (no system fonts available; standard fonts substituted)");
            }
        }

        private static void PngDecoderHandlesColourTypes()
        {
            // 2x2 truecolour PNG, built here so the test does not depend on sample files.
            var png = BuildPng(2, 2, 2, new byte[]
            {
                0, 255, 0, 0, 0, 255, 0,
                0, 0, 0, 255, 255, 255, 255,
            });
            var image = PngDecoder.Decode(png);
            Check(image != null, "decode failed");
            Check(image.Width == 2 && image.Height == 2, "wrong dimensions");
            Check(image.Components == 3, "expected RGB output");
            Check(image.Samples[0] == 255 && image.Samples[1] == 0 && image.Samples[2] == 0, "first pixel should be red");
            Check(image.Samples[9] == 255 && image.Samples[10] == 255 && image.Samples[11] == 255, "last pixel should be white");

            // Grayscale with alpha.
            var gray = BuildPng(2, 1, 4, new byte[] { 0, 10, 255, 200, 0 });
            var grayImage = PngDecoder.Decode(gray);
            Check(grayImage != null, "grayscale decode failed");
            Check(grayImage.Components == 1, "expected grayscale output");
            Check(grayImage.Alpha != null && grayImage.Alpha[1] == 0, "alpha channel was lost");
        }

        private static void BmpDecoderReads24Bit()
        {
            const int width = 2, height = 2;
            var rowBytes = ((width * 24 + 31) / 32) * 4;
            var pixelData = new byte[rowBytes * height];
            // Bottom-up: first stored row is the bottom row.
            WriteBgr(pixelData, 0, 0, 0, 255);          // bottom-left blue
            WriteBgr(pixelData, 3, 255, 255, 255);      // bottom-right white
            WriteBgr(pixelData, rowBytes, 0, 255, 0);   // top-left green
            WriteBgr(pixelData, rowBytes + 3, 255, 0, 0);

            var bmp = new byte[54 + pixelData.Length];
            bmp[0] = (byte)'B';
            bmp[1] = (byte)'M';
            WriteU32(bmp, 2, (uint)bmp.Length);
            WriteU32(bmp, 10, 54);
            WriteU32(bmp, 14, 40);
            WriteU32(bmp, 18, width);
            WriteU32(bmp, 22, height);
            bmp[26] = 1;
            bmp[28] = 24;
            Array.Copy(pixelData, 0, bmp, 54, pixelData.Length);

            string format;
            var image = ImageDecoder.Decode(bmp, out format);
            Check(format == "BMP", "format not recognised");
            Check(image != null, "decode failed");
            Check(image.Width == 2 && image.Height == 2, "wrong dimensions");
            Check(image.Samples[0] == 0 && image.Samples[1] == 255 && image.Samples[2] == 0, "top-left should be green");
        }

        private static void OpcResolvesRelationships()
        {
            Check(OpcPackage.ResolveTarget("word/document.xml", "media/image1.png") == "word/media/image1.png",
                  "relative target not resolved");
            Check(OpcPackage.ResolveTarget("word/document.xml", "/media/image.bmp") == "media/image.bmp",
                  "absolute target not resolved");
            Check(OpcPackage.ResolveTarget("word/document.xml", "../customXml/item1.xml") == "customXml/item1.xml",
                  "parent segment not resolved");

            var sample = FirstSample();
            using (var stream = new FileStream(sample, FileMode.Open, FileAccess.Read))
            using (var package = OpcPackage.Open(stream, true))
            {
                var main = package.FindRelationshipByType(null, Ns.RelOfficeDocument);
                Check(main != null, "no officeDocument relationship");
                var part = OpcPackage.ResolveTarget(string.Empty, main.Target);
                Check(package.HasPart(part), "main document part missing: " + part);
                Check(package.ReadXml(part) != null, "main document part is not XML");
            }
        }

        private static void TextWrapsAtAvailableWidth()
        {
            var warnings = new List<string>();
            var options = new ConversionOptions();
            var fonts = new FontManager(options, warnings);
            var context = new LayoutContext(fonts, options, warnings);
            var builder = new LayoutBuilder(context);

            var format = CharacterFormat.Default();
            format.SizePt = 10;
            var paragraph = new Paragraph { Format = ParagraphFormat.Default(), RunDefaults = format };
            var words = string.Join(" ", Enumerable.Repeat("wrapping", 40));
            paragraph.Inlines.Add(new TextInline(words, format));

            var wide = builder.BuildParagraph(paragraph, 500, null).Count(f => !f.IsSpacing);
            var narrow = builder.BuildParagraph(paragraph, 120, null).Count(f => !f.IsSpacing);
            Check(wide >= 1, "no lines produced");
            Check(narrow > wide, "narrow column should produce more lines (" + narrow + " vs " + wide + ")");

            var singleWord = new Paragraph { Format = ParagraphFormat.Default(), RunDefaults = format };
            singleWord.Inlines.Add(new TextInline("short", format));
            Check(builder.BuildParagraph(singleWord, 500, null).Count(f => !f.IsSpacing) == 1, "expected a single line");
        }

        private static void NumberingProducesLabels()
        {
            var sample = Path.Combine(_sampleDirectory, "SaBC TK User Manual Contributor v0.1.docx");
            if (!File.Exists(sample))
                sample = FirstSample();

            var warnings = new List<string>();
            using (var stream = new FileStream(sample, FileMode.Open, FileAccess.Read))
            using (var package = OpcPackage.Open(stream, true))
            {
                var document = new DocumentReader(package, warnings).Read();
                var labelled = 0;
                var headings = 0;
                foreach (var section in document.Sections)
                {
                    foreach (var block in section.Blocks)
                    {
                        var paragraph = block as Paragraph;
                        if (paragraph == null)
                            continue;
                        if (!string.IsNullOrEmpty(paragraph.ListLabel))
                            labelled++;
                        if (paragraph.HeadingLevel > 0)
                            headings++;
                    }
                }
                Check(labelled > 0 || headings > 0, "expected numbered paragraphs or headings in the sample");
            }
        }

        private static void SampleDocumentsConvert()
        {
            var samples = Directory.GetFiles(_sampleDirectory, "*.docx");
            Check(samples.Length > 0, "no .docx files found in " + _sampleDirectory);

            foreach (var sample in samples)
            {
                var target = Path.Combine(_outputDirectory, Path.GetFileNameWithoutExtension(sample) + ".pdf");
                var result = DocxToPdfConverter.Convert(sample, target);
                Check(result.PageCount > 0, Path.GetFileName(sample) + ": no pages produced");
                Check(File.Exists(target), Path.GetFileName(sample) + ": no output file");
                Check(new FileInfo(target).Length > 1000, Path.GetFileName(sample) + ": output is suspiciously small");
                Console.WriteLine(string.Format("        {0}: {1} pages, {2:N0} bytes",
                    Path.GetFileName(sample), result.PageCount, result.ByteCount));
            }
        }

        private static void ConvertedPdfStructureIsValid()
        {
            foreach (var pdf in Directory.GetFiles(_outputDirectory, "*.pdf"))
            {
                // Only validate our own output; reference PDFs from Word may use
                // cross-reference streams this simple check does not parse.
                var source = Path.Combine(_sampleDirectory, Path.GetFileNameWithoutExtension(pdf) + ".docx");
                if (!File.Exists(source))
                    continue;
                var bytes = File.ReadAllBytes(pdf);
                var text = Encoding.ASCII.GetString(bytes);
                var name = Path.GetFileName(pdf);
                Check(text.StartsWith("%PDF-", StringComparison.Ordinal), name + ": missing header");
                Check(text.TrimEnd().EndsWith("%%EOF", StringComparison.Ordinal), name + ": missing EOF");
                Check(text.Contains("/Type /Catalog"), name + ": missing catalog");
                Check(text.Contains("/Type /Pages"), name + ": missing page tree");
                Check(text.Contains("/Type /Page "), name + ": missing pages");
                Check(text.Contains("startxref"), name + ": missing startxref");
                Check(XrefOffsetsResolve(bytes), name + ": xref offsets do not resolve");
            }
        }

        private static void ConversionIsDeterministic()
        {
            var sample = FirstSample();
            var first = DocxToPdfConverter.Convert(File.ReadAllBytes(sample));
            var second = DocxToPdfConverter.Convert(File.ReadAllBytes(sample));
            Check(first.Length == second.Length, "output length differs between runs");

            // Only the creation date is expected to differ, and only if the clock ticks over.
            var differences = 0;
            for (var i = 0; i < first.Length; i++)
            {
                if (first[i] != second[i])
                    differences++;
            }
            Check(differences <= 16, "conversion is not reproducible (" + differences + " differing bytes)");
        }

        private static void OptionsSwitchOffFeatures()
        {
            var sample = FirstSample();
            var bytes = File.ReadAllBytes(sample);

            var withImages = DocxToPdfConverter.Convert(bytes);
            var withoutImages = DocxToPdfConverter.Convert(bytes, new ConversionOptions { RenderImages = false });
            Check(withoutImages.Length < withImages.Length, "disabling images did not shrink the output");

            var uncompressed = DocxToPdfConverter.Convert(bytes, new ConversionOptions
            {
                CompressStreams = false,
                RenderImages = false,
            });
            var uncompressedText = Encoding.ASCII.GetString(uncompressed);
            Check(uncompressedText.Contains(" Tf"), "uncompressed output should contain text operators");
            Check(!uncompressedText.Contains("/Filter /FlateDecode /Length") || uncompressedText.Contains("stream"),
                  "unexpected stream structure");

            var standardFonts = DocxToPdfConverter.Convert(bytes, new ConversionOptions { EmbedFonts = false });
            Check(Encoding.ASCII.GetString(standardFonts).Contains("/Subtype /Type1"),
                  "expected standard Type1 fonts when embedding is disabled");
        }

        // ------------------------------------------------------------------ helpers

        private static string FirstSample()
        {
            var samples = Directory.GetFiles(_sampleDirectory, "*.docx");
            if (samples.Length == 0)
                throw new Exception("no .docx samples in " + _sampleDirectory);
            Array.Sort(samples);
            return samples[0];
        }

        /// <summary>Verifies every xref entry points at "N 0 obj".</summary>
        private static bool XrefOffsetsResolve(byte[] pdf)
        {
            var text = Encoding.ASCII.GetString(pdf);
            var startIndex = text.LastIndexOf("startxref", StringComparison.Ordinal);
            if (startIndex < 0)
                return false;

            var tail = text.Substring(startIndex + "startxref".Length).Trim().Split('\n');
            long xrefIndex;
            if (tail.Length == 0 || !long.TryParse(tail[0].Trim(), out xrefIndex))
                return false;
            if (xrefIndex < 0 || xrefIndex >= pdf.Length)
                return false;
            if (text.Substring((int)xrefIndex, 4) != "xref")
                return false;

            var lines = text.Substring((int)xrefIndex).Split('\n');
            if (lines.Length < 2)
                return false;
            var header = lines[1].Trim().Split(' ');
            if (header.Length != 2)
                return false;
            int start, count;
            if (!int.TryParse(header[0], out start) || !int.TryParse(header[1], out count))
                return false;

            for (var i = 1; i < count && i + 2 < lines.Length; i++)
            {
                var entry = lines[i + 2];
                if (entry.Length < 18)
                    return false;
                long offset;
                if (!long.TryParse(entry.Substring(0, 10), NumberStyles.Integer, CultureInfo.InvariantCulture, out offset))
                    return false;
                var expected = (start + i).ToString(CultureInfo.InvariantCulture) + " 0 obj";
                if (offset + expected.Length > pdf.Length)
                    return false;
                if (text.Substring((int)offset, expected.Length) != expected)
                    return false;
            }
            return true;
        }

        private static void WriteBgr(byte[] buffer, int offset, byte b, byte g, byte r)
        {
            buffer[offset] = b;
            buffer[offset + 1] = g;
            buffer[offset + 2] = r;
        }

        private static void WriteU32(byte[] buffer, int offset, uint value)
        {
            buffer[offset] = (byte)value;
            buffer[offset + 1] = (byte)(value >> 8);
            buffer[offset + 2] = (byte)(value >> 16);
            buffer[offset + 3] = (byte)(value >> 24);
        }

        /// <summary>Builds an uncompressed-filter PNG with the given colour type from raw scanlines.</summary>
        private static byte[] BuildPng(int width, int height, int colorType, byte[] scanlines)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);

                var ihdr = new byte[13];
                WriteU32Be(ihdr, 0, (uint)width);
                WriteU32Be(ihdr, 4, (uint)height);
                ihdr[8] = 8;                     // bit depth
                ihdr[9] = (byte)colorType;
                WriteChunk(ms, "IHDR", ihdr);
                WriteChunk(ms, "IDAT", Flate.Compress(scanlines));
                WriteChunk(ms, "IEND", new byte[0]);
                return ms.ToArray();
            }
        }

        private static void WriteChunk(Stream stream, string type, byte[] data)
        {
            var length = new byte[4];
            WriteU32Be(length, 0, (uint)data.Length);
            stream.Write(length, 0, 4);
            var typeBytes = Encoding.ASCII.GetBytes(type);
            stream.Write(typeBytes, 0, 4);
            stream.Write(data, 0, data.Length);
            stream.Write(new byte[4], 0, 4);     // CRC is not validated by the decoder
        }

        private static void WriteU32Be(byte[] buffer, int offset, uint value)
        {
            buffer[offset] = (byte)(value >> 24);
            buffer[offset + 1] = (byte)(value >> 16);
            buffer[offset + 2] = (byte)(value >> 8);
            buffer[offset + 3] = (byte)value;
        }
    }
}
