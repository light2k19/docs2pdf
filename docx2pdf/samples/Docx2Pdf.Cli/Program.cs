using System;
using System.Diagnostics;
using System.IO;
using Docx2Pdf;

namespace Docx2Pdf.Cli
{
    /// <summary>Command line front end: docx2pdf &lt;input.docx&gt; [output.pdf] [options]</summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length == 0 || args[0] == "-h" || args[0] == "--help")
            {
                Usage();
                return args.Length == 0 ? 1 : 0;
            }

            string input = null;
            string output = null;
            var options = new ConversionOptions();
            var quiet = false;

            foreach (var arg in args)
            {
                if (arg.StartsWith("--", StringComparison.Ordinal))
                {
                    var parts = arg.Substring(2).Split(new[] { '=' }, 2);
                    var name = parts[0].ToLowerInvariant();
                    var value = parts.Length > 1 ? parts[1] : null;
                    switch (name)
                    {
                        case "no-embed": options.EmbedFonts = false; break;
                        case "no-images": options.RenderImages = false; break;
                        case "no-compress": options.CompressStreams = false; break;
                        case "no-outline": options.GenerateOutline = false; break;
                        case "no-links": options.CreateHyperlinks = false; break;
                        case "no-headers": options.RenderHeadersAndFooters = false; break;
                        case "quiet": quiet = true; break;
                        case "font-dir":
                            if (value != null)
                                options.FontDirectories.Add(value);
                            break;
                        case "default-font":
                            if (value != null)
                                options.DefaultFontFamily = value;
                            break;
                        case "max-pages":
                            int max;
                            if (value != null && int.TryParse(value, out max))
                                options.MaxPages = max;
                            break;
                        default:
                            Console.Error.WriteLine("Unknown option: " + arg);
                            return 2;
                    }
                }
                else if (input == null)
                {
                    input = arg;
                }
                else if (output == null)
                {
                    output = arg;
                }
            }

            if (input == null)
            {
                Usage();
                return 1;
            }
            if (output == null)
                output = Path.ChangeExtension(input, ".pdf");

            try
            {
                var stopwatch = Stopwatch.StartNew();
                var result = DocxToPdfConverter.Convert(input, output, options);
                stopwatch.Stop();

                if (!quiet)
                {
                    Console.WriteLine("{0} -> {1}", Path.GetFileName(input), Path.GetFileName(output));
                    Console.WriteLine("  pages   : {0}", result.PageCount);
                    Console.WriteLine("  size    : {0:N0} bytes", result.ByteCount);
                    Console.WriteLine("  time    : {0:N0} ms", stopwatch.ElapsedMilliseconds);
                    if (result.EmbeddedFonts.Count > 0)
                        Console.WriteLine("  fonts   : {0}", string.Join(", ", result.EmbeddedFonts));
                    foreach (var warning in result.Warnings)
                        Console.WriteLine("  warning : {0}", warning);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Conversion failed: " + ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                return 3;
            }
        }

        private static void Usage()
        {
            Console.WriteLine("docx2pdf <input.docx> [output.pdf] [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --no-embed           substitute standard PDF fonts instead of embedding font files");
            Console.WriteLine("  --no-images          skip images");
            Console.WriteLine("  --no-compress        write uncompressed streams (useful for inspecting output)");
            Console.WriteLine("  --no-outline         do not generate PDF bookmarks from headings");
            Console.WriteLine("  --no-links           do not generate link annotations");
            Console.WriteLine("  --no-headers         skip page headers and footers");
            Console.WriteLine("  --font-dir=PATH      additional directory to search for font files");
            Console.WriteLine("  --default-font=NAME  font used when the document does not name one");
            Console.WriteLine("  --max-pages=N        stop after N pages");
            Console.WriteLine("  --quiet              print nothing on success");
        }
    }
}
