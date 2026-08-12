using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using SkiaSharp;

namespace Docx2Pdf.Compare
{
    /// <summary>
    /// Compares a screenshot of a Word page with our rendered page.
    /// The two images are aligned on the page border rectangle (or the ink bounding box when the
    /// document has no page border), scaled to a common size, and then compared line by line.
    /// </summary>
    internal static class PageAlign
    {
        /// <summary>Only ink inside this horizontal band (in page points) is considered, when set.</summary>
        public static double BandLeftPt = double.NaN;
        public static double BandRightPt = double.NaN;

        /// <summary>Compares two equally sized page renders 1:1 and prints the line table.</summary>
        public static int RunIdentity(string referencePath, string oursPath, double pageHeightPt)
        {
            using (var reference = SKBitmap.Decode(referencePath))
            using (var ours = SKBitmap.Decode(oursPath))
            {
                if (reference == null || ours == null)
                {
                    Console.Error.WriteLine("Could not decode one of the images.");
                    return 1;
                }

                var scale = pageHeightPt / reference.Height;
                var referenceLines = FindLines(reference, 0, int.MaxValue);
                var oursLines = FindLines(ours, 0, int.MaxValue);
                Report(referenceLines, oursLines, scale, scale, 0, 0);
            }
            return 0;
        }

        public static int Run(string referencePath, string oursPath, string outputPath, double pageHeightPt)
        {
            using (var reference = SKBitmap.Decode(referencePath))
            using (var ours = SKBitmap.Decode(oursPath))
            {
                if (reference == null || ours == null)
                {
                    Console.Error.WriteLine("Could not decode one of the images.");
                    return 1;
                }

                // Our render is the reliable one: detect its page border first, then look for the
                // matching border in the screenshot (which may also have a window frame of its own).
                var oursRect = DetectPageRect(ours, null);
                var expected = new SKRect(
                    (float)oursRect.Left / ours.Width,
                    (float)oursRect.Top / ours.Height,
                    (float)oursRect.Right / ours.Width,
                    (float)oursRect.Bottom / ours.Height);
                var referenceRect = DetectPageRect(reference, expected);
                Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                    "reference page rect: {0},{1} {2}x{3}   ours: {4},{5} {6}x{7}",
                    referenceRect.Left, referenceRect.Top, referenceRect.Width, referenceRect.Height,
                    oursRect.Left, oursRect.Top, oursRect.Width, oursRect.Height));

                using (var referenceCrop = Crop(reference, referenceRect))
                using (var oursCrop = Crop(ours, oursRect))
                using (var oursScaled = Resize(oursCrop, referenceCrop.Width, referenceCrop.Height))
                {
                    // Our render maps 1:1 onto the page, so it defines the coordinate system;
                    // the reference screenshot is mapped onto it through the page border rectangle.
                    var oursPointsPerPixel = pageHeightPt / ours.Height;
                    var borderTopPt = oursRect.Top * oursPointsPerPixel;
                    var borderHeightPt = oursRect.Height * oursPointsPerPixel;
                    var borderLeftPt = oursRect.Left * oursPointsPerPixel;
                    var borderWidthPt = oursRect.Width * oursPointsPerPixel;

                    var scaleY = borderHeightPt / referenceCrop.Height;
                    var scaleX = borderWidthPt / referenceCrop.Width;

                    var bandLeft = double.IsNaN(BandLeftPt) ? 0 : (int)((BandLeftPt - borderLeftPt) / scaleX);
                    var bandRight = double.IsNaN(BandRightPt) ? int.MaxValue : (int)((BandRightPt - borderLeftPt) / scaleX);

                    Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                        "mapping: origin ({0:F1},{1:F1})pt  scale ({2:F4},{3:F4})pt/px  band px {4}..{5}",
                        borderLeftPt, borderTopPt, scaleX, scaleY, bandLeft, bandRight));

                    var referenceLines = FindLines(referenceCrop, bandLeft, bandRight);
                    var oursLines = FindLines(oursScaled, bandLeft, bandRight);
                    Report(referenceLines, oursLines, scaleX, scaleY, borderLeftPt, borderTopPt);

                    if (!string.IsNullOrEmpty(outputPath))
                    {
                        using (var diff = Diff(referenceCrop, oursScaled))
                            Save(diff, outputPath);
                        Console.WriteLine("diff image: " + outputPath);
                    }
                }
            }
            return 0;
        }

        private sealed class TextLine
        {
            public int Top;
            public int Bottom;
            public int Left;
            public int Right;
            public int Ink;
            public int Center { get { return (Top + Bottom) / 2; } }
        }

        /// <summary>Finds the rectangle of the printed page: the outer border, or the ink bounds.</summary>
        private static SKRectI DetectPageRect(SKBitmap bitmap, SKRect? expected)
        {
            var width = bitmap.Width;
            var height = bitmap.Height;

            // A screenshot may carry a window frame of its own, so several rectangles can qualify.
            // When the caller knows where the page border should be (from our own render), the
            // closest candidate is taken instead of the outermost one.
            const int margin = 2;

            var rowCandidates = new List<int>();
            for (var y = margin; y < height - margin; y++)
            {
                if (LongestDarkRun(bitmap, y) > width * 0.75)
                    rowCandidates.Add(y);
            }

            var columnCandidates = new List<int>();
            for (var x = margin; x < width - margin; x++)
            {
                if (LongestDarkRunColumn(bitmap, x) > height * 0.75)
                    columnCandidates.Add(x);
            }

            int borderTop = -1, borderBottom = -1, borderLeft = -1, borderRight = -1;
            if (rowCandidates.Count > 0)
            {
                borderTop = expected.HasValue ? Closest(rowCandidates, expected.Value.Top * height) : rowCandidates[0];
                borderBottom = expected.HasValue
                    ? Closest(rowCandidates, expected.Value.Bottom * height)
                    : rowCandidates[rowCandidates.Count - 1];
            }
            if (columnCandidates.Count > 0)
            {
                borderLeft = expected.HasValue ? Closest(columnCandidates, expected.Value.Left * width) : columnCandidates[0];
                borderRight = expected.HasValue
                    ? Closest(columnCandidates, expected.Value.Right * width)
                    : columnCandidates[columnCandidates.Count - 1];
            }

            if (borderTop >= 0 && borderBottom > borderTop + 50 && borderLeft >= 0 && borderRight > borderLeft + 50)
                return new SKRectI(borderLeft, borderTop, borderRight + 1, borderBottom + 1);

            // Fall back to the bounding box of everything that is not white.
            int minX = width, minY = height, maxX = 0, maxY = 0;
            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    if (Luminance(bitmap, x, y) < 200)
                    {
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
            }
            if (maxX <= minX || maxY <= minY)
                return new SKRectI(0, 0, width, height);
            return new SKRectI(minX, minY, maxX + 1, maxY + 1);
        }

        private static int Closest(List<int> candidates, double target)
        {
            var best = candidates[0];
            var bestDistance = double.MaxValue;
            foreach (var candidate in candidates)
            {
                var distance = Math.Abs(candidate - target);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    best = candidate;
                }
            }
            return best;
        }

        private static int LongestDarkRun(SKBitmap bitmap, int y)
        {
            int best = 0, run = 0;
            for (var x = 0; x < bitmap.Width; x++)
            {
                if (Luminance(bitmap, x, y) < 128)
                {
                    run++;
                    if (run > best)
                        best = run;
                }
                else
                {
                    run = 0;
                }
            }
            return best;
        }

        private static int LongestDarkRunColumn(SKBitmap bitmap, int x)
        {
            int best = 0, run = 0;
            for (var y = 0; y < bitmap.Height; y++)
            {
                if (Luminance(bitmap, x, y) < 128)
                {
                    run++;
                    if (run > best)
                        best = run;
                }
                else
                {
                    run = 0;
                }
            }
            return best;
        }

        /// <summary>Groups rows that contain ink into text lines.</summary>
        private static List<TextLine> FindLines(SKBitmap bitmap, int bandLeft, int bandRight)
        {
            var lines = new List<TextLine>();
            TextLine current = null;

            // Columns that are dark for most of the page are borders (page frame, table rules);
            // counting them would join every text line into one band.
            var isBorderColumn = new bool[bitmap.Width];
            for (var x = 0; x < bitmap.Width; x++)
            {
                // Use the ink threshold here: an anti-aliased border is not "black" but still joins lines.
                int best = 0, run = 0;
                for (var y = 0; y < bitmap.Height; y++)
                {
                    if (Luminance(bitmap, x, y) < 170)
                    {
                        run++;
                        if (run > best)
                            best = run;
                    }
                    else
                    {
                        run = 0;
                    }
                }
                isBorderColumn[x] = best > bitmap.Height * 0.5;
            }

            for (var y = 0; y < bitmap.Height; y++)
            {
                var ink = 0;
                var left = int.MaxValue;
                var right = 0;
                for (var x = Math.Max(0, bandLeft); x < Math.Min(bitmap.Width, bandRight); x++)
                {
                    if (isBorderColumn[x])
                        continue;
                    if (Luminance(bitmap, x, y) < 170)
                    {
                        ink++;
                        if (x < left) left = x;
                        if (x > right) right = x;
                    }
                }

                // Ignore the page border itself and near-empty rows.
                var isBorder = ink > bitmap.Width * 0.7;
                if (ink >= 3 && !isBorder)
                {
                    if (current == null)
                    {
                        current = new TextLine { Top = y, Bottom = y, Left = left, Right = right, Ink = ink };
                        lines.Add(current);
                    }
                    else
                    {
                        current.Bottom = y;
                        current.Ink += ink;
                        if (left < current.Left) current.Left = left;
                        if (right > current.Right) current.Right = right;
                    }
                }
                else
                {
                    current = null;
                }
            }

            return lines;
        }

        /// <summary>
        /// Pairs Word's text lines with ours by vertical position (not by index, so an extra or
        /// missing line does not shift everything) and prints the difference for each.
        /// </summary>
        private static void Report(List<TextLine> reference, List<TextLine> ours,
                                   double scaleX, double scaleY, double originX, double originY)
        {
            Func<TextLine, double> topOf = line => originY + line.Top * scaleY;
            Func<TextLine, double> bottomOf = line => originY + line.Bottom * scaleY;
            Func<TextLine, double> centerOf = line => originY + line.Center * scaleY;
            Func<TextLine, double> leftOf = line => originX + line.Left * scaleX;
            Func<TextLine, double> rightOf = line => originX + line.Right * scaleX;

            var used = new bool[ours.Count];
            var rows = new List<string>();
            double worst = 0;
            var matched = 0;
            double sum = 0;

            foreach (var a in reference)
            {
                var bestIndex = -1;
                var bestDistance = double.MaxValue;
                for (var i = 0; i < ours.Count; i++)
                {
                    if (used[i])
                        continue;
                    var distance = Math.Abs(centerOf(ours[i]) - centerOf(a));
                    // Only pair lines that also start at a similar horizontal position.
                    if (Math.Abs(leftOf(ours[i]) - leftOf(a)) > 40)
                        distance += 60;
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestIndex = i;
                    }
                }

                if (bestIndex < 0 || bestDistance > 60)
                {
                    rows.Add(string.Format(CultureInfo.InvariantCulture,
                        "{0,7:F1}..{1,-7:F1} x{2,5:F0}-{3,-5:F0} | {4,-32} | MISSING in ours",
                        topOf(a), bottomOf(a), leftOf(a), rightOf(a), string.Empty));
                    continue;
                }

                used[bestIndex] = true;
                var b = ours[bestIndex];
                var delta = centerOf(b) - centerOf(a);
                sum += Math.Abs(delta);
                matched++;
                if (Math.Abs(delta) > Math.Abs(worst))
                    worst = delta;

                rows.Add(string.Format(CultureInfo.InvariantCulture,
                    "{0,7:F1}..{1,-7:F1} x{2,5:F0}-{3,-5:F0} | {4,7:F1}..{5,-7:F1} x{6,5:F0}-{7,-5:F0} | {8,6:F1}",
                    topOf(a), bottomOf(a), leftOf(a), rightOf(a),
                    topOf(b), bottomOf(b), leftOf(b), rightOf(b), delta));
            }

            for (var i = 0; i < ours.Count; i++)
            {
                if (used[i])
                    continue;
                rows.Add(string.Format(CultureInfo.InvariantCulture,
                    "{0,-32} | {1,7:F1}..{2,-7:F1} x{3,5:F0}-{4,-5:F0} | EXTRA in ours",
                    string.Empty, topOf(ours[i]), bottomOf(ours[i]), leftOf(ours[i]), rightOf(ours[i])));
            }

            Console.WriteLine();
            Console.WriteLine("Word line (pt)                   | Ours (pt)                        | delta");
            Console.WriteLine(new string('-', 84));
            foreach (var row in rows)
                Console.WriteLine(row);

            Console.WriteLine();
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "lines: Word {0}, ours {1}, matched {2}, mean |delta| {3:F1}pt, largest {4:F1}pt",
                reference.Count, ours.Count, matched, matched == 0 ? 0 : sum / matched, worst));
        }

        private static SKBitmap Diff(SKBitmap reference, SKBitmap ours)
        {
            var result = new SKBitmap(reference.Width, reference.Height);
            for (var y = 0; y < reference.Height; y++)
            {
                for (var x = 0; x < reference.Width; x++)
                {
                    var a = Luminance(reference, x, y) < 170;
                    var b = Luminance(ours, x, y) < 170;
                    if (a && b)
                        result.SetPixel(x, y, new SKColor(60, 60, 60));
                    else if (a)
                        result.SetPixel(x, y, new SKColor(220, 40, 40));      // Word only
                    else if (b)
                        result.SetPixel(x, y, new SKColor(40, 80, 220));      // ours only
                    else
                        result.SetPixel(x, y, SKColors.White);
                }
            }
            return result;
        }

        private static SKBitmap Crop(SKBitmap source, SKRectI rect)
        {
            var result = new SKBitmap(rect.Width, rect.Height);
            using (var canvas = new SKCanvas(result))
            {
                canvas.Clear(SKColors.White);
                canvas.DrawBitmap(source, rect, new SKRect(0, 0, rect.Width, rect.Height));
            }
            return result;
        }

        private static SKBitmap Resize(SKBitmap source, int width, int height)
        {
            var result = new SKBitmap(width, height);
            using (var canvas = new SKCanvas(result))
            {
                canvas.Clear(SKColors.White);
                canvas.DrawBitmap(source, new SKRectI(0, 0, source.Width, source.Height), new SKRect(0, 0, width, height));
            }
            return result;
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

        private static byte Luminance(SKBitmap bitmap, int x, int y)
        {
            if (x < 0 || y < 0 || x >= bitmap.Width || y >= bitmap.Height)
                return 255;
            var color = bitmap.GetPixel(x, y);
            if (color.Alpha == 0)
                return 255;
            return (byte)((color.Red * 299 + color.Green * 587 + color.Blue * 114) / 1000);
        }
    }
}
