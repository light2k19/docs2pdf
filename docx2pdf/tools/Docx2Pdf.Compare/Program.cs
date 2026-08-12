using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PDFtoImage;
using SkiaSharp;

namespace Docx2Pdf.Compare
{
    /// <summary>
    /// Development tool. Rasterises a reference PDF (produced by Word) and our PDF page by page,
    /// writes the page images plus a difference image, and prints a similarity score per page.
    ///
    /// Usage: Docx2Pdf.Compare &lt;reference.pdf&gt; &lt;ours.pdf&gt; &lt;outputDirectory&gt; [dpi]
    /// </summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length >= 3 && args[0] == "--lines")
            {
                // Two same-size page renders compared 1:1 (no border alignment).
                var pageHeightForLines = args.Length > 3 ? double.Parse(args[3], CultureInfo.InvariantCulture) : 792.0;
                return PageAlign.RunIdentity(args[1], args[2], pageHeightForLines);
            }

            if (args.Length >= 3 && args[0] == "--align")
            {
                // --align <wordScreenshot.png> <ourPage.png> [diff.png] [pageHeightPt]
                var diffPath = args.Length > 3 ? args[3] : null;
                var pageHeight = args.Length > 4 ? double.Parse(args[4], CultureInfo.InvariantCulture) : 792.0;
                if (args.Length > 6)
                {
                    PageAlign.BandLeftPt = double.Parse(args[5], CultureInfo.InvariantCulture);
                    PageAlign.BandRightPt = double.Parse(args[6], CultureInfo.InvariantCulture);
                }
                return PageAlign.Run(args[1], args[2], diffPath, pageHeight);
            }

            if (args.Length == 2)
                return RenderOnly(args[0], args[1], 96, 0, int.MaxValue);
            if (args.Length == 3 && !args[2].EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)
                && File.Exists(args[0]) && !File.Exists(args[1]))
                return RenderOnly(args[0], args[1], int.Parse(args[2], CultureInfo.InvariantCulture), 0, int.MaxValue);
            if (args.Length >= 4 && args[0] == "--render")
            {
                var dpiValue = args.Length > 3 ? int.Parse(args[3], CultureInfo.InvariantCulture) : 96;
                var from = args.Length > 4 ? int.Parse(args[4], CultureInfo.InvariantCulture) : 0;
                var count = args.Length > 5 ? int.Parse(args[5], CultureInfo.InvariantCulture) : int.MaxValue;
                return RenderOnly(args[1], args[2], dpiValue, from, count);
            }

            if (args.Length < 3)
            {
                Console.Error.WriteLine("Usage: Docx2Pdf.Compare <reference.pdf> <ours.pdf> <outputDirectory> [dpi]");
                Console.Error.WriteLine("       Docx2Pdf.Compare --render <file.pdf> <outputDirectory> [dpi] [firstPage] [count]");
                return 1;
            }

            var referencePath = args[0];
            var oursPath = args[1];
            var outputDirectory = args[2];
            var dpi = args.Length > 3 ? int.Parse(args[3], CultureInfo.InvariantCulture) : 96;

            Directory.CreateDirectory(outputDirectory);

            var referenceBytes = File.ReadAllBytes(referencePath);
            var oursBytes = File.ReadAllBytes(oursPath);

            var referenceCount = Conversion.GetPageCount(referenceBytes);
            var oursCount = Conversion.GetPageCount(oursBytes);

            Console.WriteLine("reference pages: " + referenceCount);
            Console.WriteLine("ours pages     : " + oursCount);

            var pages = Math.Max(referenceCount, oursCount);
            double totalScore = 0;
            double totalInkScore = 0;
            var scored = 0;

            for (var i = 0; i < pages; i++)
            {
                var referenceImage = i < referenceCount ? Render(referenceBytes, i, dpi) : null;
                var oursImage = i < oursCount ? Render(oursBytes, i, dpi) : null;

                var name = "page-" + (i + 1).ToString("D3", CultureInfo.InvariantCulture);
                if (referenceImage != null)
                    Save(referenceImage, Path.Combine(outputDirectory, name + "-word.png"));
                if (oursImage != null)
                    Save(oursImage, Path.Combine(outputDirectory, name + "-ours.png"));

                if (referenceImage == null || oursImage == null)
                {
                    Console.WriteLine(name + ": missing on one side");
                    continue;
                }

                double score, inkScore;
                using (var diff = Difference(referenceImage, oursImage, out score, out inkScore))
                    Save(diff, Path.Combine(outputDirectory, name + "-diff.png"));

                totalScore += score;
                totalInkScore += inkScore;
                scored++;
                Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                    "{0}: {1:F1}% pixel  {2:F1}% ink  ({3}x{4} vs {5}x{6})",
                    name, score * 100, inkScore * 100, referenceImage.Width, referenceImage.Height,
                    oursImage.Width, oursImage.Height));

                referenceImage.Dispose();
                oursImage.Dispose();
            }

            if (scored > 0)
            {
                Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                    "average: {0:F1}% pixel  {1:F1}% ink  over {2} pages",
                    totalScore / scored * 100, totalInkScore / scored * 100, scored));
            }
            return 0;
        }

        /// <summary>Rasterises every page of one PDF; used to eyeball our own output.</summary>
        private static int RenderOnly(string pdfPath, string outputDirectory, int dpi, int firstPage, int count)
        {
            Directory.CreateDirectory(outputDirectory);
            var bytes = File.ReadAllBytes(pdfPath);
            var pages = Conversion.GetPageCount(bytes);
            var prefix = Path.GetFileNameWithoutExtension(pdfPath);
            var last = Math.Min(pages, firstPage + count);

            Console.WriteLine(prefix + ": " + pages + " pages");
            for (var i = firstPage; i < last; i++)
            {
                using (var bitmap = Render(bytes, i, dpi))
                {
                    var name = prefix + "-page-" + (i + 1).ToString("D3", CultureInfo.InvariantCulture) + ".png";
                    Save(bitmap, Path.Combine(outputDirectory, name));
                    Console.WriteLine("  " + name + " (" + bitmap.Width + "x" + bitmap.Height + ")");
                }
            }
            return 0;
        }

        private static SKBitmap Render(byte[] pdf, int page, int dpi)
        {
            return Conversion.ToImage(pdf, page: new Index(page), options: new RenderOptions(Dpi: dpi));
        }

        private static void Save(SKBitmap bitmap, string path)
        {
            using (var image = SKImage.FromBitmap(bitmap))
            using (var data = image.Encode(SKEncodedImageFormat.Png, 90))
            using (var stream = File.Create(path))
            {
                data.SaveTo(stream);
            }
        }

        /// <summary>
        /// Builds a red/blue difference image: red where the reference has ink we do not,
        /// blue where we have ink the reference does not.
        /// </summary>
        private static SKBitmap Difference(SKBitmap reference, SKBitmap ours,
                                           out double similarity, out double inkSimilarity)
        {
            var width = Math.Max(reference.Width, ours.Width);
            var height = Math.Max(reference.Height, ours.Height);
            var result = new SKBitmap(width, height);

            var maskA = new bool[width, height];
            var maskB = new bool[width, height];
            long same = 0;
            long total = (long)width * height;

            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    var a = Luminance(reference, x, y);
                    var b = Luminance(ours, x, y);
                    var inkA = a < 200;
                    var inkB = b < 200;
                    maskA[x, y] = inkA;
                    maskB[x, y] = inkB;

                    if (inkA == inkB)
                    {
                        same++;
                        var shade = (byte)(255 - (255 - Math.Min(a, b)) / 4);
                        result.SetPixel(x, y, new SKColor(shade, shade, shade));
                    }
                    else if (inkA)
                    {
                        result.SetPixel(x, y, new SKColor(220, 40, 40));
                    }
                    else
                    {
                        result.SetPixel(x, y, new SKColor(40, 80, 220));
                    }
                }
            }

            similarity = total == 0 ? 0 : (double)same / total;

            // Blank matching blank says nothing, and glyph strokes are 1-2px, so exact ink
            // overlap punishes harmless raster jitter. The ink score counts an ink pixel as
            // matched when the other side has ink within 2px, over all ink on either side —
            // a page of merely re-rasterised text scores high, offset text scores low.
            long inkTotal = 0;
            long inkMatched = 0;
            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    if (maskA[x, y])
                    {
                        inkTotal++;
                        if (NearInk(maskB, width, height, x, y))
                            inkMatched++;
                    }
                    if (maskB[x, y])
                    {
                        inkTotal++;
                        if (NearInk(maskA, width, height, x, y))
                            inkMatched++;
                    }
                }
            }
            inkSimilarity = inkTotal == 0 ? 1.0 : (double)inkMatched / inkTotal;
            return result;
        }

        private static bool NearInk(bool[,] mask, int width, int height, int x, int y)
        {
            var x0 = Math.Max(0, x - 2);
            var x1 = Math.Min(width - 1, x + 2);
            var y0 = Math.Max(0, y - 2);
            var y1 = Math.Min(height - 1, y + 2);
            for (var yy = y0; yy <= y1; yy++)
                for (var xx = x0; xx <= x1; xx++)
                    if (mask[xx, yy])
                        return true;
            return false;
        }

        private static byte Luminance(SKBitmap bitmap, int x, int y)
        {
            if (x >= bitmap.Width || y >= bitmap.Height)
                return 255;
            var color = bitmap.GetPixel(x, y);
            if (color.Alpha == 0)
                return 255;
            return (byte)((color.Red * 299 + color.Green * 587 + color.Blue * 114) / 1000);
        }
    }
}
