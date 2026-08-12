using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Docx2Pdf.Fonts;
using Docx2Pdf.Images;
using Docx2Pdf.Layout;
using Docx2Pdf.Model;

namespace Docx2Pdf.Pdf
{
    /// <summary>Writes laid-out pages into a PDF document.</summary>
    internal sealed class PdfRenderer
    {
        private readonly ConversionOptions _options;
        private readonly FontManager _fonts;
        private readonly PdfDocument _pdf = new PdfDocument();
        private readonly Dictionary<string, PdfReference> _imageObjects = new Dictionary<string, PdfReference>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _imageNames = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly List<PdfReference> _pageRefs = new List<PdfReference>();

        public PdfRenderer(ConversionOptions options, FontManager fonts)
        {
            _options = options;
            _fonts = fonts;
        }

        public PdfDocument Render(LayoutResult layout, DocumentInfo info)
        {
            var pagesRef = _pdf.Reserve();
            var resourcesRef = _pdf.Reserve();

            // Content streams are built first so every glyph in use is registered with its font.
            var contents = new List<byte[]>();
            var pageLinks = new List<List<LinkOp>>();
            foreach (var page in layout.Pages)
            {
                var links = new List<LinkOp>();
                contents.Add(BuildContent(page, links));
                pageLinks.Add(links);
            }

            var pageDicts = new List<PdfDictionary>();
            for (var i = 0; i < layout.Pages.Count; i++)
            {
                var page = layout.Pages[i];
                var stream = new PdfStream(_options.CompressStreams ? Flate.Compress(contents[i]) : contents[i]);
                if (_options.CompressStreams)
                    stream.Set("Filter", "FlateDecode");

                var dict = new PdfDictionary();
                dict.Set("Type", "Page");
                dict.Set("Parent", pagesRef);
                dict.Set("MediaBox", new PdfArray().Add(0).Add(0).Add(page.Width).Add(page.Height));
                dict.Set("Resources", resourcesRef);
                dict.Set("Contents", _pdf.Add(stream));
                pageDicts.Add(dict);
                _pageRefs.Add(_pdf.Add(dict));
            }

            // Link annotations can now resolve internal destinations to page objects.
            for (var i = 0; i < pageDicts.Count; i++)
            {
                var annotations = BuildAnnotations(layout.Pages[i], pageLinks[i], layout);
                if (annotations.Count == 0)
                    continue;
                var array = new PdfArray();
                foreach (var annotation in annotations)
                    array.Add(_pdf.Add(annotation));
                pageDicts[i].Set("Annots", array);
            }

            var pages = new PdfDictionary();
            pages.Set("Type", "Pages");
            pages.Set("Count", layout.Pages.Count);
            var kids = new PdfArray();
            foreach (var reference in _pageRefs)
                kids.Add(reference);
            pages.Set("Kids", kids);
            pagesRef.Target = pages;

            resourcesRef.Target = BuildResources();

            var catalog = new PdfDictionary();
            catalog.Set("Type", "Catalog");
            catalog.Set("Pages", pagesRef);

            if (_options.GenerateOutline && layout.Outline.Count > 0)
            {
                var outline = BuildOutline(layout);
                if (outline != null)
                {
                    catalog.Set("Outlines", outline);
                    catalog.Set("PageMode", "UseOutlines");
                }
            }

            _pdf.Catalog = catalog;
            _pdf.Add(catalog);
            _pdf.Info = BuildInfo(info);
            _pdf.Add(_pdf.Info);
            _pdf.FileId = (info != null ? info.Title : null) ?? "Docx2Pdf";
            return _pdf;
        }

        private PdfDictionary BuildInfo(DocumentInfo info)
        {
            var dict = new PdfDictionary();
            var title = _options.Title ?? (info != null ? info.Title : null);
            var author = _options.Author ?? (info != null ? info.Author : null);
            if (!string.IsNullOrEmpty(title))
                dict.Set("Title", new PdfString(title));
            if (!string.IsNullOrEmpty(author))
                dict.Set("Author", new PdfString(author));
            if (info != null && !string.IsNullOrEmpty(info.Subject))
                dict.Set("Subject", new PdfString(info.Subject));
            if (info != null && !string.IsNullOrEmpty(info.Keywords))
                dict.Set("Keywords", new PdfString(info.Keywords));
            dict.Set("Producer", new PdfString(_options.Producer ?? "Docx2Pdf"));
            dict.Set("Creator", new PdfString("Docx2Pdf"));
            dict.Set("CreationDate", new PdfString(FormatDate(DateTime.Now)));
            return dict;
        }

        private static string FormatDate(DateTime value)
        {
            var offset = TimeZoneInfo.Local.GetUtcOffset(value);
            var sign = offset.Ticks >= 0 ? "+" : "-";
            return "D:" + value.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
                 + sign + Math.Abs(offset.Hours).ToString("D2", CultureInfo.InvariantCulture)
                 + "'" + Math.Abs(offset.Minutes).ToString("D2", CultureInfo.InvariantCulture) + "'";
        }

        // ----------------------------------------------------------------- content streams

        private byte[] BuildContent(LayoutPage page, List<LinkOp> links)
        {
            var sb = new StringBuilder();
            var height = page.Height;
            uint currentFill = 0xFFFFFFFF;
            uint currentStroke = 0xFFFFFFFF;
            double currentLineWidth = -1;
            var dashState = string.Empty;

            // Floating frames paint above everything placed later in the flow (Word's z-order).
            var orderedOps = page.Ops;
            foreach (var candidate in page.Ops)
            {
                if (candidate.Overlay)
                {
                    orderedOps = new List<DrawOp>(page.Ops.Count);
                    foreach (var op2 in page.Ops)
                        if (!op2.Overlay)
                            orderedOps.Add(op2);
                    foreach (var op2 in page.Ops)
                        if (op2.Overlay)
                            orderedOps.Add(op2);
                    break;
                }
            }

            foreach (var op in orderedOps)
            {
                var rect = op as RectOp;
                if (rect != null)
                {
                    if (rect.Width <= 0 || rect.Height <= 0)
                        continue;
                    if (rect.Color != currentFill)
                    {
                        sb.Append(ColorOp(rect.Color, false));
                        currentFill = rect.Color;
                    }
                    sb.Append(N(rect.X)).Append(' ').Append(N(height - rect.Y - rect.Height)).Append(' ')
                      .Append(N(rect.Width)).Append(' ').Append(N(rect.Height)).Append(" re f\n");
                    continue;
                }

                var line = op as LineOp;
                if (line != null)
                {
                    if (line.Color != currentStroke)
                    {
                        sb.Append(ColorOp(line.Color, true));
                        currentStroke = line.Color;
                    }
                    var width = line.Width <= 0 ? 0.5 : line.Width;
                    if (Math.Abs(width - currentLineWidth) > 0.001)
                    {
                        sb.Append(N(width)).Append(" w\n");
                        currentLineWidth = width;
                    }

                    var dash = DashPattern(line.Style);
                    if (dash != dashState)
                    {
                        sb.Append(dash.Length == 0 ? "[] 0 d\n" : dash + "\n");
                        dashState = dash;
                    }

                    AppendLine(sb, line, height, width);
                    continue;
                }

                var text = op as TextOp;
                if (text != null)
                {
                    if (text.Font == null || string.IsNullOrEmpty(text.Text))
                        continue;
                    var hex = text.Font.EncodeHex(text.Text);
                    if (hex.Length == 0)
                        continue;

                    sb.Append("BT\n");
                    if (text.Color != currentFill)
                    {
                        sb.Append(ColorOp(text.Color, false));
                        currentFill = text.Color;
                    }
                    sb.Append('/').Append(text.Font.ResourceName).Append(' ').Append(N(text.Size)).Append(" Tf\n");
                    if (Math.Abs(text.CharSpacing) > 0.001)
                        sb.Append(N(text.CharSpacing)).Append(" Tc\n");
                    sb.Append("1 0 0 1 ").Append(N(text.X)).Append(' ').Append(N(height - text.Y)).Append(" Tm\n");
                    sb.Append('<').Append(hex).Append("> Tj\n");
                    if (Math.Abs(text.CharSpacing) > 0.001)
                        sb.Append("0 Tc\n");
                    sb.Append("ET\n");
                    continue;
                }

                var image = op as ImageOp;
                if (image != null)
                {
                    AppendImage(sb, image, height);
                    continue;
                }

                var link = op as LinkOp;
                if (link != null)
                    links.Add(link);
            }

            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static void AppendLine(StringBuilder sb, LineOp line, double pageHeight, double width)
        {
            var y1 = pageHeight - line.Y1;
            var y2 = pageHeight - line.Y2;

            if (line.Style == BorderStyle.Double)
            {
                var offset = Math.Max(0.6, width);
                var horizontal = Math.Abs(line.Y1 - line.Y2) < 0.01;
                if (horizontal)
                {
                    Segment(sb, line.X1, y1 + offset, line.X2, y2 + offset);
                    Segment(sb, line.X1, y1 - offset, line.X2, y2 - offset);
                }
                else
                {
                    Segment(sb, line.X1 - offset, y1, line.X2 - offset, y2);
                    Segment(sb, line.X1 + offset, y1, line.X2 + offset, y2);
                }
                return;
            }

            Segment(sb, line.X1, y1, line.X2, y2);
        }

        private static void Segment(StringBuilder sb, double x1, double y1, double x2, double y2)
        {
            sb.Append(N(x1)).Append(' ').Append(N(y1)).Append(" m ")
              .Append(N(x2)).Append(' ').Append(N(y2)).Append(" l S\n");
        }

        private static string DashPattern(BorderStyle style)
        {
            switch (style)
            {
                case BorderStyle.Dotted: return "[1 2] 0 d";
                case BorderStyle.Dashed: return "[4 2] 0 d";
                default: return string.Empty;
            }
        }

        private void AppendImage(StringBuilder sb, ImageOp image, double pageHeight)
        {
            if (image.Image == null || image.Width <= 0 || image.Height <= 0)
                return;

            var name = RegisterImage(image);
            if (name == null)
                return;

            var x = image.X;
            var y = pageHeight - image.Y - image.Height;

            sb.Append("q\n");
            var rotation = NormalizeAngle(image.RotationDeg);
            if (Math.Abs(rotation) > 0.01)
            {
                // Rotate about the centre of the placement rectangle.
                var cx = x + image.Width / 2;
                var cy = y + image.Height / 2;
                var radians = -rotation * Math.PI / 180.0;
                var cos = Math.Cos(radians);
                var sin = Math.Sin(radians);
                sb.Append("1 0 0 1 ").Append(N(cx)).Append(' ').Append(N(cy)).Append(" cm\n");
                sb.Append(N(cos)).Append(' ').Append(N(sin)).Append(' ')
                  .Append(N(-sin)).Append(' ').Append(N(cos)).Append(" 0 0 cm\n");
                sb.Append("1 0 0 1 ").Append(N(-cx)).Append(' ').Append(N(-cy)).Append(" cm\n");
            }
            var cropLeft = Clamp01(image.CropLeft);
            var cropRight = Clamp01(image.CropRight);
            var cropTop = Clamp01(image.CropTop);
            var cropBottom = Clamp01(image.CropBottom);
            var visibleW = 1 - cropLeft - cropRight;
            var visibleH = 1 - cropTop - cropBottom;
            if (visibleW < 0.999 || visibleH < 0.999)
            {
                if (visibleW < 0.02) { cropLeft = cropRight = 0; visibleW = 1; }
                if (visibleH < 0.02) { cropTop = cropBottom = 0; visibleH = 1; }
                // The full bitmap is scaled up and clipped so only the un-cropped part shows.
                sb.Append(N(x)).Append(' ').Append(N(y)).Append(' ')
                  .Append(N(image.Width)).Append(' ').Append(N(image.Height)).Append(" re W n\n");
                var fullW = image.Width / visibleW;
                var fullH = image.Height / visibleH;
                sb.Append(N(fullW)).Append(" 0 0 ").Append(N(fullH)).Append(' ')
                  .Append(N(x - cropLeft * fullW)).Append(' ').Append(N(y - cropBottom * fullH)).Append(" cm\n");
            }
            else
            {
                sb.Append(N(image.Width)).Append(" 0 0 ").Append(N(image.Height)).Append(' ')
                  .Append(N(x)).Append(' ').Append(N(y)).Append(" cm\n");
            }
            sb.Append('/').Append(name).Append(" Do\nQ\n");
        }

        private static double Clamp01(double value)
        {
            return value < 0 ? 0 : (value > 0.98 ? 0.98 : value);
        }

        private static double NormalizeAngle(double degrees)
        {
            var value = degrees % 360.0;
            if (value > 180) value -= 360;
            if (value < -180) value += 360;
            return value;
        }

        private string RegisterImage(ImageOp op)
        {
            var key = op.Key ?? ("img" + _imageObjects.Count.ToString(CultureInfo.InvariantCulture));
            string name;
            if (_imageNames.TryGetValue(key, out name))
                return name;

            var reference = BuildImageObject(op.Image);
            if (reference == null)
                return null;

            name = "Im" + (_imageObjects.Count + 1).ToString(CultureInfo.InvariantCulture);
            _imageObjects[name] = reference;
            _imageNames[key] = name;
            return name;
        }

        private PdfReference BuildImageObject(DecodedImage image)
        {
            if (image == null)
                return null;

            PdfStream stream;
            if (image.IsJpeg)
            {
                stream = new PdfStream(image.JpegData);
                stream.Set("Filter", "DCTDecode");
                stream.Set("ColorSpace", image.JpegComponents == 1 ? "DeviceGray"
                                       : image.JpegComponents == 4 ? "DeviceCMYK" : "DeviceRGB");
                if (image.JpegComponents == 4 && image.JpegAdobeInverted)
                {
                    var decode = new PdfArray();
                    for (var i = 0; i < 4; i++)
                    {
                        decode.Add(1);
                        decode.Add(0);
                    }
                    stream.Set("Decode", decode);
                }
            }
            else
            {
                if (image.Samples == null)
                    return null;
                stream = BuildSampleStream(image.Samples, image.Width, image.Height,
                                           image.Components == 1 ? 1 : 3);
                stream.Set("ColorSpace", image.Components == 1 ? "DeviceGray" : "DeviceRGB");
            }

            stream.Set("Type", "XObject");
            stream.Set("Subtype", "Image");
            stream.Set("Width", image.Width);
            stream.Set("Height", image.Height);
            stream.Set("BitsPerComponent", 8);

            if (image.Alpha != null)
            {
                var mask = BuildSampleStream(image.Alpha, image.Width, image.Height, 1);
                mask.Set("Type", "XObject");
                mask.Set("Subtype", "Image");
                mask.Set("Width", image.Width);
                mask.Set("Height", image.Height);
                mask.Set("ColorSpace", "DeviceGray");
                mask.Set("BitsPerComponent", 8);
                stream.Set("SMask", _pdf.Add(mask));
            }

            return _pdf.Add(stream);
        }

        /// <summary>
        /// Image samples deflate ~2x better after PNG row prediction (each row filtered
        /// against the previous one, like PNG itself does; /Predictor 15 = per-row tags).
        /// </summary>
        private PdfStream BuildSampleStream(byte[] samples, int width, int height, int colors)
        {
            var rowLen = width * colors;
            if (!_options.CompressStreams || rowLen <= 0 || (long)rowLen * height != samples.Length)
            {
                var plain = new PdfStream(_options.CompressStreams ? Flate.Compress(samples) : samples);
                if (_options.CompressStreams)
                    plain.Set("Filter", "FlateDecode");
                return plain;
            }

            var filtered = new byte[(rowLen + 1) * height];
            var previous = new byte[rowLen];
            var candidate = new byte[rowLen];
            var best = new byte[rowLen];
            for (var y = 0; y < height; y++)
            {
                var row = y * rowLen;
                var bestScore = long.MaxValue;
                byte bestTag = 0;
                for (byte tag = 0; tag <= 4; tag++)
                {
                    long score = 0;
                    for (var i = 0; i < rowLen; i++)
                    {
                        var raw = samples[row + i];
                        int left = i >= colors ? samples[row + i - colors] : 0;
                        int up = previous[i];
                        int upLeft = i >= colors ? previous[i - colors] : 0;
                        int predicted;
                        switch (tag)
                        {
                            case 1: predicted = left; break;
                            case 2: predicted = up; break;
                            case 3: predicted = (left + up) / 2; break;
                            case 4:
                                var p = left + up - upLeft;
                                var pa = Math.Abs(p - left);
                                var pb = Math.Abs(p - up);
                                var pc = Math.Abs(p - upLeft);
                                predicted = pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft;
                                break;
                            default: predicted = 0; break;
                        }
                        var value = (byte)(raw - predicted);
                        candidate[i] = value;
                        score += value < 128 ? value : 256 - value;
                    }
                    if (score < bestScore)
                    {
                        bestScore = score;
                        bestTag = tag;
                        var swap = best; best = candidate; candidate = swap;
                    }
                }
                filtered[y * (rowLen + 1)] = bestTag;
                Array.Copy(best, 0, filtered, y * (rowLen + 1) + 1, rowLen);
                Array.Copy(samples, row, previous, 0, rowLen);
            }

            var stream = new PdfStream(Flate.Compress(filtered));
            stream.Set("Filter", "FlateDecode");
            var parms = new PdfDictionary();
            parms.Set("Predictor", 15);
            parms.Set("Colors", colors);
            parms.Set("BitsPerComponent", 8);
            parms.Set("Columns", width);
            stream.Set("DecodeParms", parms);
            return stream;
        }

        private PdfDictionary BuildResources()
        {
            var resources = new PdfDictionary();

            var fonts = new PdfDictionary();
            foreach (var font in _fonts.Fonts)
            {
                if (font.ResourceName == null)
                    continue;
                fonts.Set(font.ResourceName, font.Build(_pdf));
            }
            resources.Set("Font", fonts);

            if (_imageObjects.Count > 0)
            {
                var xobjects = new PdfDictionary();
                foreach (var entry in _imageObjects)
                    xobjects.Set(entry.Key, entry.Value);
                resources.Set("XObject", xobjects);
            }

            resources.Set("ProcSet", new PdfArray(new PdfName("PDF"), new PdfName("Text"),
                                                  new PdfName("ImageB"), new PdfName("ImageC")));
            return resources;
        }

        // ----------------------------------------------------------------- annotations and outline

        private List<PdfDictionary> BuildAnnotations(LayoutPage page, List<LinkOp> links, LayoutResult layout)
        {
            var result = new List<PdfDictionary>();
            if (!_options.CreateHyperlinks)
                return result;

            foreach (var link in links)
            {
                var annotation = new PdfDictionary();
                annotation.Set("Type", "Annot");
                annotation.Set("Subtype", "Link");
                annotation.Set("Rect", new PdfArray()
                    .Add(link.X).Add(page.Height - link.Y - link.Height)
                    .Add(link.X + link.Width).Add(page.Height - link.Y));
                annotation.Set("Border", new PdfArray().Add(0).Add(0).Add(0));
                annotation.Set("F", 4);

                if (!string.IsNullOrEmpty(link.Url))
                {
                    var action = new PdfDictionary();
                    action.Set("Type", "Action");
                    action.Set("S", "URI");
                    action.Set("URI", new PdfString(link.Url));
                    annotation.Set("A", action);
                }
                else if (!string.IsNullOrEmpty(link.Anchor))
                {
                    KeyValuePair<int, double> destination;
                    if (!layout.Bookmarks.TryGetValue(link.Anchor, out destination))
                        continue;
                    annotation.Set("Dest", Destination(destination.Key, destination.Value, layout));
                }
                else
                {
                    continue;
                }

                result.Add(annotation);
            }
            return result;
        }

        private PdfArray Destination(int pageIndex, double y, LayoutResult layout)
        {
            if (pageIndex < 0 || pageIndex >= _pageRefs.Count)
                pageIndex = 0;
            var height = layout.Pages[pageIndex].Height;
            var array = new PdfArray();
            array.Add(_pageRefs[pageIndex]);
            array.Add(new PdfName("XYZ"));
            array.Add(0);
            array.Add(Math.Max(0, height - y + 12));
            array.Add(PdfNull.Instance);
            return array;
        }

        private PdfObject BuildOutline(LayoutResult layout)
        {
            // Build a tree from the flat, ordered list of headings.
            var roots = new List<OutlineEntry>();
            var stack = new List<OutlineEntry>();
            foreach (var entry in layout.Outline)
            {
                while (stack.Count > 0 && stack[stack.Count - 1].Level >= entry.Level)
                    stack.RemoveAt(stack.Count - 1);
                if (stack.Count == 0)
                    roots.Add(entry);
                else
                    stack[stack.Count - 1].Children.Add(entry);
                stack.Add(entry);
            }
            if (roots.Count == 0)
                return null;

            var rootDict = new PdfDictionary();
            rootDict.Set("Type", "Outlines");
            var rootRef = _pdf.Add(rootDict);

            int count;
            var children = BuildOutlineLevel(roots, rootRef, layout, out count);
            if (children.Key == null)
                return null;
            rootDict.Set("First", children.Key);
            rootDict.Set("Last", children.Value);
            rootDict.Set("Count", count);
            return rootRef;
        }

        private KeyValuePair<PdfReference, PdfReference> BuildOutlineLevel(
            List<OutlineEntry> entries, PdfReference parent, LayoutResult layout, out int visibleCount)
        {
            visibleCount = 0;
            var refs = new List<PdfReference>();
            var dicts = new List<PdfDictionary>();

            foreach (var entry in entries)
            {
                var dict = new PdfDictionary();
                dicts.Add(dict);
                refs.Add(_pdf.Add(dict));
            }

            for (var i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                var dict = dicts[i];
                dict.Set("Title", new PdfString(Shorten(entry.Title)));
                dict.Set("Parent", parent);
                if (i > 0)
                    dict.Set("Prev", refs[i - 1]);
                if (i < entries.Count - 1)
                    dict.Set("Next", refs[i + 1]);
                dict.Set("Dest", Destination(entry.PageIndex, entry.Y, layout));
                visibleCount++;

                if (entry.Children.Count > 0)
                {
                    int childCount;
                    var children = BuildOutlineLevel(entry.Children, refs[i], layout, out childCount);
                    if (children.Key != null)
                    {
                        dict.Set("First", children.Key);
                        dict.Set("Last", children.Value);
                        dict.Set("Count", -childCount);      // collapsed by default
                    }
                }
            }

            return refs.Count == 0
                ? new KeyValuePair<PdfReference, PdfReference>(null, null)
                : new KeyValuePair<PdfReference, PdfReference>(refs[0], refs[refs.Count - 1]);
        }

        private static string Shorten(string title)
        {
            if (string.IsNullOrEmpty(title))
                return " ";
            title = title.Replace("\r", " ").Replace("\n", " ").Trim();
            return title.Length > 160 ? title.Substring(0, 157) + "..." : title;
        }

        private static string ColorOp(uint color, bool stroke)
        {
            var r = ((color >> 16) & 0xFF) / 255.0;
            var g = ((color >> 8) & 0xFF) / 255.0;
            var b = (color & 0xFF) / 255.0;
            return N(r) + " " + N(g) + " " + N(b) + (stroke ? " RG\n" : " rg\n");
        }

        private static string N(double value)
        {
            return PdfNumber.Format(value);
        }
    }
}
