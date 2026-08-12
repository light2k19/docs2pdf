using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Docx2Pdf.Fonts;
using Docx2Pdf.Layout;
using Docx2Pdf.Model;
using Docx2Pdf.Ooxml;

namespace Docx2Pdf.Tests
{
    /// <summary>Prints how a document was interpreted: sections, tables, and per-page fragment usage.</summary>
    internal static class Diagnostics
    {
        private static string Describe(Block block)
        {
            var paragraph = block as Paragraph;
            if (paragraph != null)
            {
                var sb = new StringBuilder();
                if (!string.IsNullOrEmpty(paragraph.ListLabel))
                    sb.Append('[').Append(paragraph.ListLabel).Append("] ");
                foreach (var inline in paragraph.Inlines)
                {
                    var text = inline as TextInline;
                    if (text != null)
                        sb.Append(text.Text);
                    else if (inline is ImageInline)
                        sb.Append("<image ").Append(((ImageInline)inline).WidthPt.ToString("F0"))
                          .Append('x').Append(((ImageInline)inline).HeightPt.ToString("F0")).Append('>');
                    else if (inline is BreakInline)
                        sb.Append("<br>");
                }
                var content = sb.ToString().Trim();
                var style = paragraph.StyleId ?? "-";
                return "P(" + style + ") " + (content.Length > 60 ? content.Substring(0, 60) + "..." : content);
            }

            var table = block as Table;
            return table != null ? "TABLE rows=" + table.Rows.Count : block.GetType().Name;
        }

        /// <summary>Prints the first and last words of text on every laid-out page.</summary>
        public static void DumpPageStarts(string docxPath, TextWriter output)
        {
            var warnings = new List<string>();
            using (var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read))
            using (var package = OpcPackage.Open(stream, true))
            {
                var document = new DocumentReader(package, warnings).Read();
                var options = new ConversionOptions();
                var fonts = new FontManager(options, warnings);
                var context = new LayoutContext(fonts, options, warnings);
                var layout = new DocumentLayout(context).Layout(document);

                for (var i = 0; i < layout.Pages.Count; i++)
                {
                    var page = layout.Pages[i];
                    // Body text only: skip header/footer bands (top and bottom ~90pt).
                    var texts = new List<KeyValuePair<double, string>>();
                    foreach (var op in page.Ops)
                    {
                        var text = op as TextOp;
                        if (text == null || string.IsNullOrWhiteSpace(text.Text))
                            continue;
                        if (text.Y < 95 || text.Y > page.Height - 60)
                            continue;
                        texts.Add(new KeyValuePair<double, string>(text.Y, text.Text));
                    }
                    texts.Sort((a, b) => a.Key.CompareTo(b.Key));

                    var first = string.Empty;
                    var seenY = -1.0;
                    foreach (var t in texts)
                    {
                        if (seenY >= 0 && t.Key - seenY > 0.5 && first.Length > 45)
                            break;
                        first += t.Value + " ";
                        seenY = t.Key;
                        if (first.Length > 90)
                            break;
                    }

                    var last = string.Empty;
                    for (var j = texts.Count - 1; j >= 0 && last.Length < 45; j--)
                        last = texts[j].Value + " " + last;

                    output.WriteLine("PAGE {0}|{1}|{2}", i + 1, Clean(first, 70), Clean(last, 45));
                }
            }
        }

        private static string Clean(string text, int max)
        {
            text = System.Text.RegularExpressions.Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim();
            return text.Length > max ? text.Substring(0, max) : text;
        }

        public static void Dump(string docxPath, TextWriter output)
        {
            var warnings = new List<string>();
            using (var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read))
            using (var package = OpcPackage.Open(stream, true))
            {
                var document = new DocumentReader(package, warnings).Read();
                output.WriteLine("=== " + Path.GetFileName(docxPath) + " ===");
                output.WriteLine("sections: " + document.Sections.Count);

                for (var s = 0; s < document.Sections.Count; s++)
                {
                    var section = document.Sections[s];
                    var props = section.Properties;
                    output.WriteLine(string.Format(
                        "  section {0}: page {1:F1}x{2:F1}pt margins L{3:F1} R{4:F1} T{5:F1} B{6:F1} content {7:F1}x{8:F1} blocks {9}",
                        s, props.PageWidthPt, props.PageHeightPt, props.MarginLeftPt, props.MarginRightPt,
                        props.MarginTopPt, props.MarginBottomPt, props.ContentWidthPt, props.ContentHeightPt,
                        section.Blocks.Count));

                    var measureOptions = new ConversionOptions();
                    var measureFonts = new FontManager(measureOptions, warnings);
                    var measureContext = new LayoutContext(measureFonts, measureOptions, warnings);
                    var measure = new LayoutBuilder(measureContext) { MaxBlockHeight = props.PageHeightPt / 2 };
                    Func<HeaderFooter, double> heightOf = part =>
                    {
                        if (part == null)
                            return 0;
                        double total = 0;
                        foreach (var fragment in measure.BuildBlocks(part.Blocks, props.ContentWidthPt))
                            total += fragment.Height;
                        return total;
                    };
                    if (props.HeaderDefault != null)
                    {
                        var headerIndex = 0;
                        double headerRunning = 0;
                        foreach (var block in props.HeaderDefault.Blocks)
                        {
                            var label = Describe(block);
                            foreach (var fragment in measure.BuildBlocks(new[] { block }, props.ContentWidthPt))
                            {
                                headerRunning += fragment.Height;
                                output.WriteLine(string.Format("    hdr {0,3}: h={1,6:F1} total={2,7:F1} {3}{4}",
                                    headerIndex++, fragment.Height, headerRunning,
                                    fragment.IsSpacing ? "[spacing] " : string.Empty, label));
                                label = string.Empty;
                            }
                        }
                    }

                    var headerHeight = heightOf(props.HeaderDefault);
                    var footerHeight = heightOf(props.FooterDefault);
                    var top = Math.Max(props.MarginTopPt, headerHeight > 0 ? props.HeaderDistancePt + headerHeight : 0);
                    var bottom = Math.Max(props.MarginBottomPt, footerHeight > 0 ? props.FooterDistancePt + footerHeight : 0);
                    output.WriteLine(string.Format(
                        "    header {0:F1}pt (distance {1:F1}) footer {2:F1}pt (distance {3:F1}) => body top {4:F1}, height {5:F1}",
                        headerHeight, props.HeaderDistancePt, footerHeight, props.FooterDistancePt,
                        top, props.PageHeightPt - top - bottom));

                    var tableIndex = 0;
                    foreach (var block in section.Blocks)
                    {
                        var table = block as Table;
                        if (table == null)
                            continue;
                        var grid = new StringBuilder();
                        foreach (var w in table.Grid)
                            grid.Append(w.ToString("F1")).Append(' ');
                        output.WriteLine(string.Format("    table {0}: rows {1} grid[{2}] = {3} pref {4}{5}",
                            tableIndex++, table.Rows.Count, table.Grid.Count, grid.ToString().Trim(),
                            table.PreferredWidthPt.HasValue ? table.PreferredWidthPt.Value.ToString("F1") : "auto",
                            table.PreferredWidthIsPercent ? "%" : "pt"));
                    }
                }

                var options = new ConversionOptions();
                var fonts = new FontManager(options, warnings);
                var context = new LayoutContext(fonts, options, warnings);

                // Fragment-by-fragment budget for the first section, to explain page breaks.
                var probe = new LayoutBuilder(context) { MaxBlockHeight = document.Sections[0].Properties.ContentHeightPt };
                double running = 0;
                var index = 0;
                foreach (var block in document.Sections[0].Blocks)
                {
                    var text = Describe(block);
                    foreach (var fragment in probe.BuildBlocks(new[] { block }, document.Sections[0].Properties.ContentWidthPt))
                    {
                        running += fragment.Height;
                        output.WriteLine(string.Format("    frag {0,3}: h={1,6:F1} total={2,7:F1} {3}{4}",
                            index++, fragment.Height, running, fragment.IsSpacing ? "[spacing] " : string.Empty, text));
                        text = string.Empty;
                    }
                }

                var layout = new DocumentLayout(context).Layout(document);

                output.WriteLine("pages: " + layout.Pages.Count);
                for (var i = 0; i < layout.Pages.Count; i++)
                {
                    var page = layout.Pages[i];
                    int texts = 0, rects = 0, lines = 0, images = 0, links = 0;
                    double minY = double.MaxValue, maxY = double.MinValue;
                    foreach (var op in page.Ops)
                    {
                        var text = op as TextOp;
                        if (text != null)
                        {
                            texts++;
                            minY = Math.Min(minY, text.Y);
                            maxY = Math.Max(maxY, text.Y);
                            continue;
                        }
                        if (op is RectOp) rects++;
                        else if (op is LineOp) lines++;
                        else if (op is ImageOp) images++;
                        else if (op is LinkOp) links++;
                    }
                    output.WriteLine(string.Format(
                        "  page {0}: text {1} rect {2} line {3} image {4} link {5} textY {6:F0}..{7:F0}",
                        i + 1, texts, rects, lines, images, links,
                        texts == 0 ? 0 : minY, texts == 0 ? 0 : maxY));
                }

                output.WriteLine("outline entries: " + layout.Outline.Count);
                foreach (var entry in layout.Outline)
                    output.WriteLine("  outline L" + entry.Level + " p" + (entry.PageIndex + 1) + " " + entry.Title);
                foreach (var warning in warnings)
                    output.WriteLine("warning: " + warning);
            }
        }
    }
}
