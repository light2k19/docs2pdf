using System;
using System.Collections.Generic;
using System.Text;
using Docx2Pdf.Fonts;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    /// <summary>Turns document blocks into a flat list of vertically stacked fragments.</summary>
    internal sealed partial class LayoutBuilder
    {
        private readonly LayoutContext _ctx;

        /// <summary>True while building the content of a table cell (affects auto spacing).</summary>
        private bool _inTableCell;

        public LayoutBuilder(LayoutContext context)
        {
            _ctx = context;
        }

        public List<Fragment> BuildBlocks(IEnumerable<Block> blocks, double availableWidth)
        {
            var list = blocks as IList<Block> ?? new List<Block>(blocks);
            var fragments = new List<Fragment>();
            Paragraph previous = null;
            var previousWasTable = false;
            for (var i = 0; i < list.Count; i++)
            {
                var paragraph = list[i] as Paragraph;
                if (paragraph != null)
                {
                    var next = i + 1 < list.Count ? list[i + 1] as Paragraph : null;
                    var nextIsTable = i + 1 < list.Count && list[i + 1] is Table;
                    fragments.AddRange(BuildParagraph(paragraph, availableWidth, previous, next,
                                                      previousWasTable, nextIsTable,
                                                      i == 0, i == list.Count - 1));
                    previous = paragraph;
                    previousWasTable = false;
                    continue;
                }

                var table = list[i] as Table;
                if (table != null)
                {
                    fragments.AddRange(BuildTable(table, availableWidth));
                    previous = null;
                    previousWasTable = true;
                }
            }
            return fragments;
        }

        // ----------------------------------------------------------------- atoms

        private abstract class Atom
        {
            public double Width;
            public double Ascent;
            public double Descent;      // positive distance below the baseline
            public double LineHeight;
            public bool IsSpace;
            public string LinkUrl;
            public string LinkAnchor;
        }

        private sealed class TextAtom : Atom
        {
            public PdfFontBase Font;
            public string Text;
            public double Size;
            public uint Color;
            public double BaselineShift;
            public UnderlineStyle Underline = UnderlineStyle.None;
            public uint UnderlineColor;
            public bool Strike;
            public uint? Highlight;
            public double CharSpacing;
            public FieldKind Field;
        }

        private sealed class ImageAtom : Atom
        {
            public Model.ImageInline Source;
            public Images.DecodedImage Image;
            /// <summary>Size of the picture itself, excluding reserved effect space.</summary>
            public double DrawWidth;
            public double DrawHeight;
            /// <summary>Line height of the run's font, for the leading a spacing multiple adds.</summary>
            public double RunLineHeight;
        }

        private sealed class TabAtom : Atom
        {
            public TabLeader Leader = TabLeader.None;
            public PdfFontBase Font;
            public double Size;
            public uint Color;
            /// <summary>Fixed target position, used for the tab that follows a list label.</summary>
            public double? TargetX;
        }

        private sealed class BreakAtom : Atom
        {
            public BreakKind Kind;
        }

        private sealed class BookmarkAtom : Atom
        {
            public string Name;
        }

        /// <summary>A floating frame; it occupies no space in the line.</summary>
        private sealed class AnchorAtom : Atom
        {
            public AnchoredInline Source;
            public List<DrawOp> Content;
            public double FrameHeight;
        }

        // ----------------------------------------------------------------- paragraphs

        public List<Fragment> BuildParagraph(Paragraph paragraph, double availableWidth, Paragraph previous)
        {
            return BuildParagraph(paragraph, availableWidth, previous, null, false, false, false, false);
        }

        public List<Fragment> BuildParagraph(Paragraph paragraph, double availableWidth,
                                             Paragraph previous, Paragraph next,
                                             bool previousIsTable, bool nextIsTable,
                                             bool firstInContainer, bool lastInContainer)
        {
            var fragments = new List<Fragment>();
            var format = paragraph.Format ?? new ParagraphFormat();

            // Word honours negative indents by letting text run into the margin.
            var indentLeft = format.IndentLeftPt ?? 0;
            var indentRight = format.IndentRightPt ?? 0;
            var firstLineIndent = format.IndentFirstLinePt ?? 0;
            if (availableWidth - indentLeft - indentRight < 12)
            {
                indentLeft = 0;
                indentRight = 0;
            }

            // HTML "auto" spacing (beforeAutospacing/afterAutospacing) resolves to Word's
            // classic 14pt HTML margin, and adjacent margins collapse like HTML margins:
            // the boundary gets the larger of the two contributions, not their sum.
            // The paragraph's own after-spacer always carries its full contribution; the
            // before-spacer carries only what the previous paragraph's after did not cover.
            // HTML "auto" spacing (w:beforeAutospacing / w:afterAutospacing), calibrated
            // directly against Word's rendering of a probe document:
            //  - an auto margin contributes Word's classic 14pt, everywhere (body, cells,
            //    against paragraphs, tables and image paragraphs alike);
            //  - adjacent margins collapse HTML-style: the boundary gets the larger
            //    contribution, not the sum;
            //  - an auto margin at the start or end of a table cell collapses to nothing;
            //  - contextualSpacing (same style adjacency) overrides everything to zero.
            //  - between two consecutive items of the same list (same w:numId) an auto margin
            //    contributes nothing, the way HTML list items carry no vertical margin inside
            //    their list; the list's outer boundaries still get the full 14pt
            //    (PROBE6: X->B1 28.5pt, B1->B2 14.7pt, B3->X 28.8pt).
            const double htmlAutoSpacePt = 14.0;
            var sameListAsPrevious = paragraph.ListNumId != null && previous != null
                && string.Equals(previous.ListNumId, paragraph.ListNumId, StringComparison.Ordinal);
            var sameListAsNext = paragraph.ListNumId != null && next != null
                && string.Equals(next.ListNumId, paragraph.ListNumId, StringComparison.Ordinal);
            var spaceBefore = format.AutoSpaceBefore == true
                ? (sameListAsPrevious ? 0 : htmlAutoSpacePt)
                : (format.SpaceBeforePt ?? 0);
            var spaceAfter = format.AutoSpaceAfter == true
                ? (sameListAsNext ? 0 : htmlAutoSpacePt)
                : (format.SpaceAfterPt ?? 0);

            if (_inTableCell && firstInContainer && format.AutoSpaceBefore == true)
                spaceBefore = 0;
            if (_inTableCell && lastInContainer && format.AutoSpaceAfter == true)
                spaceAfter = 0;

            // contextualSpacing ("don't add space between paragraphs of the same style"):
            // the paragraph's own space-before vanishes when the previous paragraph has the same
            // style, and its space-after vanishes when the next one does.
            if (format.ContextualSpacing == true && previous != null
                && string.Equals(previous.StyleId, paragraph.StyleId, StringComparison.OrdinalIgnoreCase))
            {
                spaceBefore = 0;
            }
            if (format.ContextualSpacing == true && next != null
                && string.Equals(next.StyleId, paragraph.StyleId, StringComparison.OrdinalIgnoreCase))
            {
                spaceAfter = 0;
            }

            // Margin collapsing with the previous paragraph: its after-spacer was already
            // emitted at full value, so this paragraph's before-spacer only adds what the
            // previous contribution did not cover.
            var previousFormat = previous != null ? previous.Format : null;
            var previousAutoAfter = previousFormat != null && previousFormat.AutoSpaceAfter == true;
            if (format.AutoSpaceBefore == true || previousAutoAfter)
            {
                double previousAfter = 0;
                if (previousFormat != null)
                {
                    previousAfter = previousAutoAfter
                        ? (sameListAsPrevious ? 0 : htmlAutoSpacePt)
                        : (previousFormat.SpaceAfterPt ?? 0);
                    if (previousFormat.ContextualSpacing == true
                        && string.Equals(previous.StyleId, paragraph.StyleId, StringComparison.OrdinalIgnoreCase))
                        previousAfter = 0;
                }
                spaceBefore = Math.Max(0, Math.Max(spaceBefore, previousAfter) - previousAfter);
            }

            if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_SPACE") != null && (spaceBefore > 0 || spaceAfter > 0))
            {
                var probe0 = PlainText(paragraph);
                Console.Error.WriteLine("space before={0:F1} after={1:F1} autoB={2} autoA={3} num={4} '{5}'",
                    spaceBefore, spaceAfter, format.AutoSpaceBefore, format.AutoSpaceAfter,
                    paragraph.ListNumId ?? "-", probe0.Length > 30 ? probe0.Substring(0, 30) : probe0);
            }

            var atoms = BuildAtoms(paragraph);
            var lines = BreakIntoLines(paragraph, atoms, availableWidth, indentLeft, indentRight, firstLineIndent);

            var paragraphId = new object();
            var widowControl = format.WidowControl != false;

            if (spaceBefore > 0)
            {
                var spacer = new Fragment(spaceBefore)
                {
                    IsSpacing = true,
                    IsSpaceBefore = true,
                    PageBreakBefore = format.PageBreakBefore == true,
                    ParagraphId = paragraphId,
                };
                if (format.Shading.HasValue)
                    spacer.Add(new RectOp { X = 0, Y = 0, Width = availableWidth, Height = spaceBefore, Color = format.Shading.Value });
                fragments.Add(spacer);
            }

            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                // Consecutive paragraphs with identical borders form one bordered group in
                // Word: the top edge draws only above the first of the group and the bottom
                // only below the last (the Xu portfolio precedes each underlined section
                // heading with an empty paragraph carrying the same pBdr — one rule, not two).
                var borderTop = previous == null || previous.Format == null
                                || !SameBorders(format.Borders, previous.Format.Borders);
                var borderBottom = next == null || next.Format == null
                                   || !SameBorders(format.Borders, next.Format.Borders);
                var fragment = EmitLine(paragraph, line, availableWidth, i == 0, i == lines.Count - 1,
                                        borderTop, borderBottom);
                if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_PARA") != null)
                {
                    var probe = PlainText(paragraph);
                    Console.Error.WriteLine(
                        "para line{0}: h={1:F2} w={2:F2}/{3:F2} asc={4:F2} fontH={5:F2} mult={6:F2} '{7}'",
                        i, fragment.Height, line.ContentWidth, line.MaxWidth, line.Ascent, line.FontHeight,
                        format.LineSpacing ?? 1.0,
                        probe.Length > 40 ? probe.Substring(0, 40) : probe);
                }
                fragment.ParagraphId = paragraphId;
                fragment.LineIndex = i;
                fragment.LineCount = lines.Count;
                fragment.WidowControl = widowControl;
                if (i == 0 && fragments.Count == 0)
                    fragment.PageBreakBefore = format.PageBreakBefore == true;
                if (i == 0 && paragraph.HeadingLevel > 0)
                {
                    fragment.HeadingLevel = paragraph.HeadingLevel;
                    fragment.HeadingText = PlainText(paragraph);
                }
                // An explicit page break (w:br w:type="page") breaks even when it is the very
                // first thing in the paragraph; the composer ignores it at the top of a page.
                if (line.PageBreakBefore)
                    fragment.PageBreakBefore = true;
                if (line.PageBreakAfter)
                    fragment.PageBreakAfter = true;
                fragment.KeepWithNext = format.KeepLines == true && i < lines.Count - 1;
                fragments.Add(fragment);
            }

            if (spaceAfter > 0)
            {
                var spacer = new Fragment(spaceAfter) { IsSpacing = true, ParagraphId = paragraphId };
                if (format.Shading.HasValue)
                    spacer.Add(new RectOp { X = 0, Y = 0, Width = availableWidth, Height = spaceAfter, Color = format.Shading.Value });
                fragments.Add(spacer);
            }

            // Keep-with-next applies to the whole paragraph: the decision must be made before
            // its first fragment is placed, or the heading is orphaned at the page bottom.
            if (format.KeepNext == true)
            {
                foreach (var fragment in fragments)
                    fragment.KeepWithNext = true;
            }

            return fragments;
        }

        private static string PlainText(Paragraph paragraph)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(paragraph.ListLabel))
                sb.Append(paragraph.ListLabel).Append(' ');
            foreach (var inline in paragraph.Inlines)
            {
                var text = inline as TextInline;
                if (text != null)
                    sb.Append(text.Text);
                else if (inline is TabInline)
                    sb.Append(' ');
            }
            return sb.ToString().Trim();
        }

        private List<Atom> BuildAtoms(Paragraph paragraph)
        {
            var atoms = new List<Atom>();

            if (!string.IsNullOrEmpty(paragraph.ListLabel))
            {
                var labelFormat = paragraph.ListLabelFormat ?? paragraph.RunDefaults;
                AppendTextAtoms(atoms, paragraph.ListLabel, labelFormat, null, null, true);
                if (paragraph.ListLabelFollowedByTab)
                {
                    var labelFont = _ctx.ResolveFont(labelFormat);
                    var labelSize = SizeOf(labelFormat);
                    atoms.Add(new TabAtom
                    {
                        Font = labelFont,
                        Size = labelSize,
                        Color = ColorOf(labelFormat),
                        TargetX = paragraph.ListTextIndentPt,
                        Ascent = labelFont.AscentEm * labelSize,
                        Descent = -labelFont.DescentEm * labelSize,
                        LineHeight = labelFont.LineHeightEm * labelSize,
                    });
                }
            }

            foreach (var inline in paragraph.Inlines)
            {
                var format = inline.Format ?? paragraph.RunDefaults;
                if (format != null && format.Hidden == true)
                    continue;

                var text = inline as TextInline;
                if (text != null)
                {
                    AppendTextAtoms(atoms, text.Text, format, inline.LinkUrl, inline.LinkAnchor, false);
                    continue;
                }

                var field = inline as FieldInline;
                if (field != null)
                {
                    var placeholder = field.Kind == FieldKind.NumPages ? "99" : "1";
                    var start = atoms.Count;
                    AppendTextAtoms(atoms, placeholder, format, inline.LinkUrl, inline.LinkAnchor, true);
                    for (var i = start; i < atoms.Count; i++)
                    {
                        var atom = atoms[i] as TextAtom;
                        if (atom != null)
                            atom.Field = field.Kind;
                    }
                    continue;
                }

                var tab = inline as TabInline;
                if (tab != null)
                {
                    var font = _ctx.ResolveFont(format);
                    var size = SizeOf(format);
                    atoms.Add(new TabAtom
                    {
                        Font = font,
                        Size = size,
                        Color = ColorOf(format),
                        Ascent = font.AscentEm * size,
                        Descent = -font.DescentEm * size,
                        LineHeight = font.LineHeightEm * size,
                        LinkUrl = inline.LinkUrl,
                        LinkAnchor = inline.LinkAnchor,
                    });
                    continue;
                }

                var brk = inline as BreakInline;
                if (brk != null)
                {
                    var font = _ctx.ResolveFont(format);
                    var size = SizeOf(format);
                    atoms.Add(new BreakAtom
                    {
                        Kind = brk.Kind,
                        Ascent = font.AscentEm * size,
                        Descent = -font.DescentEm * size,
                        LineHeight = font.LineHeightEm * size,
                    });
                    continue;
                }

                var image = inline as ImageInline;
                if (image != null)
                {
                    AppendImageAtom(atoms, image, format);
                    continue;
                }

                var anchored = inline as AnchoredInline;
                if (anchored != null)
                {
                    AppendAnchorAtom(atoms, anchored);
                    continue;
                }

                var bookmark = inline as BookmarkInline;
                if (bookmark != null)
                    atoms.Add(new BookmarkAtom { Name = bookmark.Name });
            }

            return atoms;
        }

        /// <summary>Lays a floating frame's content out once; it is positioned later by the composer.</summary>
        private void AppendAnchorAtom(List<Atom> atoms, AnchoredInline anchored)
        {
            if (!_ctx.Options.RenderImages && anchored.Blocks.Count == 1)
            {
                var only = anchored.Blocks[0] as Paragraph;
                if (only != null && only.Inlines.Count == 1 && only.Inlines[0] is ImageInline)
                    return;
            }

            var width = anchored.WidthPt > 1 ? anchored.WidthPt : 200;
            var previousMax = MaxBlockHeight;
            MaxBlockHeight = 100000;              // a frame is never split across pages
            List<Fragment> fragments;
            try
            {
                fragments = BuildBlocks(anchored.Blocks, width);
            }
            finally
            {
                MaxBlockHeight = previousMax;
            }

            var ops = new List<DrawOp>();
            var frameW = anchored.WidthPt;
            var frameH = anchored.HeightPt;
            if (anchored.FillColor.HasValue && frameW > 0 && frameH > 0)
                ops.Add(new RectOp { X = 0, Y = 0, Width = frameW, Height = frameH, Color = anchored.FillColor.Value });

            double y = 0;
            foreach (var fragment in fragments)
            {
                foreach (var op in fragment.Ops)
                {
                    if (op is BookmarkOp)
                        continue;
                    ops.Add(op.Translate(0, y));
                }
                y += fragment.Height;
            }

            if (anchored.OutlineColor.HasValue && frameW > 0 && frameH > 0)
            {
                var color = anchored.OutlineColor.Value;
                var lineWidth = anchored.OutlineWidthPt;
                ops.Add(new LineOp { X1 = 0, Y1 = 0, X2 = frameW, Y2 = 0, Width = lineWidth, Color = color });
                ops.Add(new LineOp { X1 = 0, Y1 = frameH, X2 = frameW, Y2 = frameH, Width = lineWidth, Color = color });
                ops.Add(new LineOp { X1 = 0, Y1 = 0, X2 = 0, Y2 = frameH, Width = lineWidth, Color = color });
                ops.Add(new LineOp { X1 = frameW, Y1 = 0, X2 = frameW, Y2 = frameH, Width = lineWidth, Color = color });
            }

            if (ops.Count == 0)
                return;

            atoms.Add(new AnchorAtom
            {
                Source = anchored,
                Content = ops,
                FrameHeight = anchored.HeightPt > 0 ? anchored.HeightPt : y,
                Width = 0,
                Ascent = 0,
                Descent = 0,
                LineHeight = 0,
            });
        }

        private void AppendImageAtom(List<Atom> atoms, ImageInline image, CharacterFormat format)
        {
            if (!_ctx.Options.RenderImages)
                return;

            var decoded = _ctx.GetImage(image);
            var width = image.WidthPt;
            var height = image.HeightPt;

            if (decoded != null)
            {
                if (width <= 0 && height <= 0)
                {
                    // No explicit size: assume 96 DPI.
                    width = decoded.Width * 72.0 / 96.0;
                    height = decoded.Height * 72.0 / 96.0;
                }
                else if (width <= 0)
                {
                    width = height * decoded.AspectRatio;
                }
                else if (height <= 0)
                {
                    height = width / decoded.AspectRatio;
                }
            }
            if (width <= 0 || height <= 0)
                return;

            var font = _ctx.ResolveFont(format);
            var size = SizeOf(format);

            atoms.Add(new ImageAtom
            {
                Source = image,
                Image = decoded,
                DrawWidth = width,
                DrawHeight = height,
                // Shadow and other effects reserve additional layout space around the picture
                // on every side — verified by probe against Word's rendering.
                Width = width + image.EffectLeftPt + image.EffectRightPt,
                Ascent = height + image.EffectTopPt + image.EffectBottomPt,
                Descent = 0,
                LineHeight = height + image.EffectTopPt + image.EffectBottomPt,
                // The run's font height feeds the leading a line-spacing multiple adds
                // below the picture (PROBE11, V4: the run's font, not the mark's).
                RunLineHeight = font.LineHeightEm * size,
                LinkUrl = image.LinkUrl,
                LinkAnchor = image.LinkAnchor,
            });
        }

        private double SizeOf(CharacterFormat format)
        {
            var size = format != null && format.SizePt.HasValue ? format.SizePt.Value : 11;
            if (size <= 0)
                size = 11;
            if (format != null && format.VertAlign.HasValue && format.VertAlign.Value != VerticalTextAlignment.Baseline)
                size *= 0.65;
            return size;
        }

        private static uint ColorOf(CharacterFormat format)
        {
            return format != null && format.Color.HasValue ? format.Color.Value : 0x000000u;
        }

        private void AppendTextAtoms(List<Atom> atoms, string text, CharacterFormat format,
                                     string linkUrl, string linkAnchor, bool noBreakAfter)
        {
            if (string.IsNullOrEmpty(text))
                return;

            if (format != null && format.AllCaps == true)
                text = text.ToUpperInvariant();

            var smallCaps = format != null && format.SmallCaps == true;
            var baseSize = SizeOf(format);
            var bold = format != null && format.Bold == true;
            var italic = format != null && format.Italic == true;
            var color = ColorOf(format);
            var primary = _ctx.ResolveFont(format);
            var baselineShift = 0.0;
            if (format != null && format.VertAlign.HasValue)
            {
                if (format.VertAlign.Value == VerticalTextAlignment.Superscript)
                    baselineShift = -0.45 * baseSize;
                else if (format.VertAlign.Value == VerticalTextAlignment.Subscript)
                    baselineShift = 0.15 * baseSize;
            }

            foreach (var token in Tokenize(text))
            {
                var tokenText = token.Value;
                var isSpace = token.Key;

                foreach (var piece in SplitSmallCaps(tokenText, smallCaps))
                {
                    var size = piece.Key ? baseSize * 0.8 : baseSize;
                    var pieceText = piece.Value;
                    foreach (var run in _ctx.Fonts.Split(pieceText, primary, bold, italic))
                    {
                        var font = run.Font;
                        var charSpacing = format != null && format.CharacterSpacingPt.HasValue ? format.CharacterSpacingPt.Value : 0;
                        var width = font.Measure(run.Text, size) + charSpacing * run.Text.Length;
                        atoms.Add(new TextAtom
                        {
                            Font = font,
                            Text = run.Text,
                            Size = size,
                            Color = color,
                            Width = width,
                            Ascent = font.AscentEm * size - baselineShift,
                            Descent = -font.DescentEm * size + baselineShift,
                            LineHeight = font.LineHeightEm * baseSize,
                            IsSpace = isSpace,
                            BaselineShift = baselineShift,
                            Underline = format != null && format.Underline.HasValue ? format.Underline.Value : UnderlineStyle.None,
                            UnderlineColor = format != null && format.UnderlineColor.HasValue ? format.UnderlineColor.Value : color,
                            Strike = format != null && format.Strike == true,
                            Highlight = format != null ? (format.Highlight ?? format.Shading) : null,
                            CharSpacing = charSpacing,
                            LinkUrl = linkUrl,
                            LinkAnchor = linkAnchor,
                        });
                    }
                }
            }
        }

        /// <summary>Splits text into break-separated tokens; the key marks whitespace tokens.</summary>
        private static IEnumerable<KeyValuePair<bool, string>> Tokenize(string text)
        {
            var i = 0;
            while (i < text.Length)
            {
                var c = text[i];
                if (IsSpaceChar(c))
                {
                    var start = i;
                    while (i < text.Length && (IsSpaceChar(text[i])))
                        i++;
                    yield return new KeyValuePair<bool, string>(true, text.Substring(start, i - start));
                }
                else if (IsWideBreakable(c))
                {
                    yield return new KeyValuePair<bool, string>(false, text.Substring(i, 1));
                    i++;
                }
                else
                {
                    var start = i;
                    while (i < text.Length)
                    {
                        var ch = text[i];
                        if (IsSpaceChar(ch) || IsWideBreakable(ch))
                            break;
                        i++;
                        // Allow a break directly after a hyphen or slash.
                        if (ch == '-' || ch == '/' || ch == '–' || ch == '—')
                            break;
                    }
                    yield return new KeyValuePair<bool, string>(false, text.Substring(start, i - start));
                }
            }
        }

        /// <summary>True when a line may break directly after this character (Tokenize splits there).</summary>
        internal static bool AllowsBreakAfter(char c)
        {
            return c == '-' || c == '/' || c == '–' || c == '—' || IsWideBreakable(c);
        }

        private static bool IsSpaceChar(char c)
        {
            // space, tab, no-break space, en/em/thin spaces and zero-width space
            return c == ' ' || c == '\t' || c == '\u00A0' || c == '\u2002'
                || c == '\u2003' || c == '\u2009' || c == '\u200B' || c == '\u3000';
        }

        /// <summary>CJK and other scripts that may break between any two characters.</summary>
        private static bool IsWideBreakable(char c)
        {
            return (c >= 0x1100 && c <= 0x11FF)
                || (c >= 0x2E80 && c <= 0x303E)
                || (c >= 0x3041 && c <= 0x33FF)
                || (c >= 0x3400 && c <= 0x4DBF)
                || (c >= 0x4E00 && c <= 0x9FFF)
                || (c >= 0xA000 && c <= 0xA4CF)
                || (c >= 0xAC00 && c <= 0xD7AF)
                || (c >= 0xF900 && c <= 0xFAFF)
                || (c >= 0xFF00 && c <= 0xFF60)
                || (c >= 0xFFE0 && c <= 0xFFE6);
        }

        /// <summary>Splits into runs of "was lowercase" (rendered smaller) when small caps are active.</summary>
        private static IEnumerable<KeyValuePair<bool, string>> SplitSmallCaps(string text, bool smallCaps)
        {
            if (!smallCaps || text.Length == 0)
            {
                yield return new KeyValuePair<bool, string>(false, text);
                yield break;
            }

            var start = 0;
            var currentLower = char.IsLower(text[0]);
            for (var i = 1; i <= text.Length; i++)
            {
                var lower = i < text.Length && char.IsLower(text[i]);
                if (i == text.Length || lower != currentLower)
                {
                    var slice = text.Substring(start, i - start);
                    yield return new KeyValuePair<bool, string>(currentLower, currentLower ? slice.ToUpperInvariant() : slice);
                    start = i;
                    currentLower = lower;
                }
            }
        }
    }
}
