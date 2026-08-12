using System;
using System.Collections.Generic;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    internal sealed partial class LayoutBuilder
    {
        private sealed class Line
        {
            public readonly List<Atom> Atoms = new List<Atom>();
            public double Left;
            public double MaxWidth;
            public double Width;              // including trailing whitespace
            public double ContentWidth;       // excluding trailing whitespace
            public bool EndsWithBreak;
            public bool PageBreakBefore;
            public bool PageBreakAfter;
            public bool HasTab;
            public double Ascent;
            public double Descent;
            public double NaturalHeight;
            /// <summary>Tallest font-based line height; line-spacing multiples apply to this only.</summary>
            public double FontHeight;
            /// <summary>Tallest inline picture box (extent + effect space) on the line.</summary>
            public double ImageBox;
            /// <summary>Font height of the tallest picture's run, for the multiple's leading.</summary>
            public double ImageRunFontHeight;

            public bool HasContent
            {
                get
                {
                    foreach (var atom in Atoms)
                    {
                        if (!atom.IsSpace && !(atom is BookmarkAtom))
                            return true;
                    }
                    return false;
                }
            }

            /// <summary>Total width of inter-word spaces on the line (justification compression).</summary>
            public double SpaceWidth;

            public void Add(Atom atom)
            {
                Atoms.Add(atom);
                Width += atom.Width;
                if (atom.IsSpace)
                    SpaceWidth += atom.Width;
                if (!atom.IsSpace)
                    ContentWidth = Width;
                if (atom.Ascent > Ascent) Ascent = atom.Ascent;
                if (atom.Descent > Descent) Descent = atom.Descent;
                if (atom.LineHeight > NaturalHeight) NaturalHeight = atom.LineHeight;
                // Images and floating frames are excluded: Word applies proportional line
                // spacing to text metrics, not to inline object heights. A picture instead
                // records its box and its run's font height — the spacing multiple's extra
                // leading, (multiple − 1) × font height, hangs below the picture (PROBE11/12).
                var image = atom as ImageAtom;
                if (image != null)
                {
                    if (atom.LineHeight > ImageBox)
                    {
                        ImageBox = atom.LineHeight;
                        ImageRunFontHeight = image.RunLineHeight;
                    }
                }
                else if (!(atom is AnchorAtom) && atom.LineHeight > FontHeight)
                    FontHeight = atom.LineHeight;
            }
        }

        private List<Line> BreakIntoLines(Paragraph paragraph, List<Atom> atoms, double availableWidth,
                                          double indentLeft, double indentRight, double firstLineIndent)
        {
            var lines = new List<Line>();
            var right = Math.Max(indentLeft + 12, availableWidth - indentRight);

            // Empty lines take the height of the paragraph mark, which may be formatted
            // differently from the paragraph's text (w:pPr/w:rPr).
            var markFormat = paragraph.MarkFormat ?? paragraph.RunDefaults;
            var defaultFont = _ctx.ResolveFont(markFormat);
            var defaultSize = SizeOf(markFormat);

            // Lines beside a pending floating table lose the space it occupies; once the
            // float's height is consumed, following lines regain the full width (Word
            // wraps mid-paragraph: sample1 p3 has six narrow lines beside the ITEM
            // table, then continues full-width below it).
            var floatState = !_inTableCell ? _pendingFloat : null;
            var floatRemaining = floatState != null ? floatState.RemainingPt : 0;

            Func<bool, Line> newLine = first =>
            {
                var left = first ? indentLeft + firstLineIndent : indentLeft;
                var lineRight = right;
                if (floatState != null && floatRemaining > 0.5)
                {
                    left += floatState.OccupiedPt;
                    if (floatState.OccupiedRightPt > 0)
                        lineRight = Math.Max(left + 12, right - floatState.OccupiedRightPt);
                }
                return new Line
                {
                    Left = left,
                    MaxWidth = Math.Max(1, lineRight - left),
                    Ascent = 0,
                    Descent = 0,
                    NaturalHeight = 0,
                };
            };

            var format0 = paragraph.Format ?? new ParagraphFormat();
            Action<Line> commit = line =>
            {
                lines.Add(line);
                // Anchor-only lines collapse to zero flow height, so they must not
                // consume the float's vertical extent either.
                if (floatState != null && !AnchorOnlyLine(line))
                    floatRemaining -= LineBoxHeight(format0, line);
            };

            var current = newLine(true);
            var pageBreakPending = false;

            for (var i = 0; i < atoms.Count; i++)
            {
                var atom = atoms[i];

                var brk = atom as BreakAtom;
                if (brk != null)
                {
                    if (brk.Kind == BreakKind.Line)
                    {
                        EnsureMetrics(current, defaultFont, defaultSize);
                        current.EndsWithBreak = true;
                        current.PageBreakBefore |= pageBreakPending;
                        pageBreakPending = false;
                        commit(current);
                        current = newLine(false);
                    }
                    else
                    {
                        if (current.Atoms.Count > 0)
                        {
                            EnsureMetrics(current, defaultFont, defaultSize);
                            current.EndsWithBreak = true;
                            current.PageBreakBefore |= pageBreakPending;
                            pageBreakPending = false;
                            commit(current);
                            current = newLine(false);
                        }
                        pageBreakPending = true;
                    }
                    continue;
                }

                if (atom is BookmarkAtom)
                {
                    current.Add(atom);
                    continue;
                }

                var tab = atom as TabAtom;
                if (tab != null)
                {
                    var currentX = current.Left + current.Width;
                    var target = ResolveTabStop(paragraph, tab, atoms, i, currentX, right);
                    if (target <= currentX + 0.01)
                    {
                        // No stop left on this line: wrap.
                        if (current.HasContent)
                        {
                            EnsureMetrics(current, defaultFont, defaultSize);
                            current.PageBreakBefore |= pageBreakPending;
                            pageBreakPending = false;
                            commit(current);
                            current = newLine(false);
                        }
                        continue;
                    }
                    tab.Width = target - currentX;
                    current.HasTab = true;
                    current.Add(tab);
                    continue;
                }

                // Justified space compression is a MODE-15 behaviour: Word 2013+'s engine
                // squeezes inter-word spaces to pull a word up (SaBC, mode 15 — the 20%
                // model was calibrated there), while the legacy engine breaks rigidly
                // (sample-6, mode 14: line 1 + "augue" = 487.8pt in a 481.7pt column =
                // a 15.9%-per-space squeeze Word refuses; rigid fitting took the doc
                // 75.9 → 99.7). Left-aligned spaces are always rigid — testing without
                // the pending space admitted an extra word on razor lines (MyLesen p10:
                // "…Fi Pentadbiran." 200.6pt kept in a 199.3pt column).
                var justified = (paragraph.Format.Alignment ?? TextAlignment.Left) == TextAlignment.Justify;
                // LEGACY Word admits hairline overshoots — sample1 p5's drop-cap line
                // fits at 0.28pt over our arithmetic (sub-point space advances over ten
                // spaces) — while rejecting real ones (sample-6: 6.2pt over breaks).
                // Mode-15 razors cut the other way (a universal 0.4 dropped SaBC 2.6,
                // MyLesen 0.5): their fit stays at the pixel epsilon, with the
                // calibrated space-compression model for justified text.
                var tolerance = _ctx.LegacyCellSpacing ? 0.4 : 0.01;
                var basis = current.Width;
                if (justified && !_ctx.LegacyCellSpacing)
                {
                    tolerance += current.SpaceWidth * 0.20;
                    basis = current.ContentWidth;
                }
                if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_WRAP") != null)
                {
                    var ta = atom as TextAtom;
                    if (ta != null && !ta.IsSpace && basis + atom.Width > current.MaxWidth + tolerance)
                        Console.Error.WriteLine("wrap: reject '{0}' w={1:F2} basis={2:F2} max={3:F2} tol={4:F2} left={5:F2}",
                            ta.Text, atom.Width, basis, current.MaxWidth, tolerance, current.Left);
                }
                var fits = basis + atom.Width <= current.MaxWidth + tolerance;
                if (!fits && current.HasContent && !atom.IsSpace)
                {
                    // Word breaks lines only at whitespace. A word assembled from several runs
                    // (a bold fragment plus bare trailing punctuation, split quotes) wraps as
                    // one glued cluster instead of stranding its tail at the line end.
                    List<Atom> carried = null;
                    // The cluster glue joins the pieces of one WORD: it applies only when
                    // the overflowing atom is text. An inline image (or tab/anchor) breaks
                    // freely after text — carrying the preceding word with it shifted the
                    // picture right and stranded "application." under it (MyLesen p14).
                    var carryFrom = atom is TextAtom ? current.Atoms.Count : 0;
                    while (carryFrom > 0)
                    {
                        var prev = current.Atoms[carryFrom - 1];
                        if (prev.IsSpace || prev is TabAtom || prev is ImageAtom || prev is AnchorAtom)
                            break;
                        // Tokenize splits after hyphens, slashes and CJK characters — those
                        // are legitimate break points, so the cluster glue must not carry
                        // them back together (MyLesen p5: Word ends the line at "re-" and
                        // wraps "testing"; gluing moved the whole word and added a line).
                        var prevText = prev as TextAtom;
                        if (prevText != null && !string.IsNullOrEmpty(prevText.Text)
                            && AllowsBreakAfter(prevText.Text[prevText.Text.Length - 1]))
                            break;
                        carryFrom--;
                    }
                    if (carryFrom > 0 && carryFrom < current.Atoms.Count)
                    {
                        var kept = current.Atoms.GetRange(0, carryFrom);
                        var keptHasContent = false;
                        foreach (var keptAtom in kept)
                        {
                            if (!keptAtom.IsSpace && !(keptAtom is BookmarkAtom))
                            {
                                keptHasContent = true;
                                break;
                            }
                        }
                        if (keptHasContent)
                        {
                            carried = current.Atoms.GetRange(carryFrom, current.Atoms.Count - carryFrom);
                            var rebuilt = newLine(lines.Count == 0);
                            rebuilt.HasTab = current.HasTab;
                            foreach (var keptAtom in kept)
                                rebuilt.Add(keptAtom);
                            current = rebuilt;
                        }
                    }

                    EnsureMetrics(current, defaultFont, defaultSize);
                    current.PageBreakBefore |= pageBreakPending;
                    pageBreakPending = false;
                    commit(current);
                    current = newLine(false);
                    if (carried != null)
                    {
                        foreach (var carriedAtom in carried)
                            current.Add(carriedAtom);
                    }
                }

                if (atom.IsSpace && !current.HasContent && lines.Count > 0 && current.Atoms.Count == 0)
                    continue;      // swallow the space that caused the wrap

                // A single word wider than an EMPTY line wraps at CHARACTER granularity,
                // the way Word breaks the calendar's "11" into "1"/"1" inside a 10.8pt
                // day cell (43.2pt column − 10.8 margins − 21.6 first-line indent).
                if (!atom.IsSpace && !current.HasContent && current.Atoms.Count == 0
                    && atom.Width > current.MaxWidth - current.Width + tolerance)
                {
                    var wide = atom as TextAtom;
                    if (wide != null && wide.Text != null && wide.Text.Length > 1)
                    {
                        var limit = current.MaxWidth - current.Width + tolerance;
                        var cut = 1;
                        for (var n = wide.Text.Length - 1; n > 1; n--)
                        {
                            if (wide.Font.Measure(wide.Text.Substring(0, n), wide.Size) + wide.CharSpacing * n <= limit)
                            {
                                cut = n;
                                break;
                            }
                        }
                        var headText = wide.Text.Substring(0, cut);
                        var tailText = wide.Text.Substring(cut);
                        var headWidth = wide.Font.Measure(headText, wide.Size) + wide.CharSpacing * headText.Length;
                        var tailWidth = wide.Font.Measure(tailText, wide.Size) + wide.CharSpacing * tailText.Length;
                        current.Add(wide.WithText(headText, headWidth));
                        EnsureMetrics(current, defaultFont, defaultSize);
                        current.PageBreakBefore |= pageBreakPending;
                        pageBreakPending = false;
                        commit(current);
                        current = newLine(false);
                        atoms[i] = wide.WithText(tailText, tailWidth);
                        i--;
                        continue;
                    }
                }

                current.Add(atom);
            }

            EnsureMetrics(current, defaultFont, defaultSize);
            if (pageBreakPending && !current.HasContent)
            {
                // The paragraph ends with a page break: Word keeps this (empty) remainder line
                // and the paragraph mark on the current page and breaks after it.
                current.PageBreakAfter = true;
            }
            else
            {
                current.PageBreakBefore |= pageBreakPending;
            }
            commit(current);
            if (floatState != null)
                floatState.RemainingPt = Math.Max(0, floatRemaining);
            return lines;
        }

        /// <summary>True when the line holds only floating anchors (it takes no flow height).</summary>
        private static bool AnchorOnlyLine(Line line)
        {
            var hasAnchor = false;
            foreach (var atom in line.Atoms)
            {
                if (atom is AnchorAtom)
                    hasAnchor = true;
                else if (!(atom is BookmarkAtom) && !atom.IsSpace)
                    return false;
            }
            return hasAnchor;
        }

        /// <summary>
        /// The box height EmitLine will give this line — used to consume a floating
        /// table's vertical extent while breaking lines (keep in sync with EmitLine's
        /// height computation).
        /// </summary>
        private static double LineBoxHeight(ParagraphFormat format, Line line)
        {
            var natural = line.Ascent + line.Descent;
            var rule = format.LineSpacingRule ?? LineSpacingRule.Auto;
            var spacing = format.LineSpacing ?? 1.0;
            if (rule == LineSpacingRule.Exact && spacing > 0)
                return spacing;
            if (rule == LineSpacingRule.AtLeast && spacing > 0)
                return Math.Max(spacing, Math.Max(line.NaturalHeight, natural));
            var multiple = spacing <= 0 ? 1.0 : spacing;
            // A sub-single AUTO multiple compresses below the glyph height (sample3:
            // w:line=120 spacer paragraphs render 6.7pt in Word, half the font line).
            var height = multiple < 1
                ? Math.Max(natural, line.FontHeight) * multiple
                : Math.Max(natural, line.FontHeight * multiple);
            if (line.ImageBox > 0)
                height = Math.Max(height, line.ImageBox + Math.Max(0, multiple - 1) * line.ImageRunFontHeight);
            return height;
        }

        private static void EnsureMetrics(Line line, Fonts.PdfFontBase font, double size)
        {
            if (line.Ascent <= 0 && line.Descent <= 0)
            {
                line.Ascent = font.AscentEm * size;
                line.Descent = -font.DescentEm * size;
                line.NaturalHeight = font.LineHeightEm * size;
                // An empty line is a font-metrics line: the line-spacing multiple applies to it.
                line.FontHeight = line.NaturalHeight;
            }
        }

        /// <summary>Finds the tab stop a tab character advances to, honouring centre/right/decimal stops.</summary>
        private double ResolveTabStop(Paragraph paragraph, TabAtom tab, List<Atom> atoms, int index,
                                      double currentX, double right)
        {
            if (tab.TargetX.HasValue && tab.TargetX.Value > currentX)
                return Math.Min(tab.TargetX.Value, right);

            TabStop stop = null;
            var tabs = paragraph.Format != null ? paragraph.Format.Tabs : null;
            if (tabs != null)
            {
                foreach (var candidate in tabs)
                {
                    if (candidate.Alignment == TabAlignment.Clear)
                        continue;
                    if (candidate.PositionPt > currentX + 0.01 && (stop == null || candidate.PositionPt < stop.PositionPt))
                        stop = candidate;
                }
            }

            double target;
            if (stop != null)
            {
                target = stop.PositionPt;
                tab.Leader = stop.Leader;
                if (stop.Alignment == TabAlignment.Center || stop.Alignment == TabAlignment.Right
                    || stop.Alignment == TabAlignment.Decimal)
                {
                    var segment = MeasureSegment(atoms, index + 1, stop.Alignment == TabAlignment.Decimal);
                    var start = stop.Alignment == TabAlignment.Center ? target - segment / 2 : target - segment;
                    target = Math.Max(currentX, start);
                }
            }
            else
            {
                var step = _ctx.DefaultTabStopPt > 0 ? _ctx.DefaultTabStopPt : 36;
                target = (Math.Floor(currentX / step) + 1) * step;
            }

            return Math.Min(target, right);
        }

        /// <summary>Width of the atoms following a tab, up to the next tab or line break.</summary>
        private static double MeasureSegment(List<Atom> atoms, int start, bool stopAtDecimal)
        {
            double width = 0;
            for (var i = start; i < atoms.Count; i++)
            {
                var atom = atoms[i];
                if (atom is TabAtom || atom is BreakAtom)
                    break;
                if (stopAtDecimal)
                {
                    var text = atom as TextAtom;
                    if (text != null && text.Text != null && text.Text.IndexOf('.') >= 0)
                    {
                        var upto = text.Text.Substring(0, text.Text.IndexOf('.'));
                        width += text.Font.Measure(upto, text.Size);
                        break;
                    }
                }
                width += atom.Width;
            }
            return width;
        }

        // ----------------------------------------------------------------- emission

        private Fragment EmitLine(Paragraph paragraph, Line line, double availableWidth, bool isFirst, bool isLast,
                                  bool borderTop = true, bool borderBottom = true)
        {
            var format = paragraph.Format ?? new ParagraphFormat();

            // Vertical metrics.
            var natural = line.Ascent + line.Descent;
            var box = Math.Max(line.NaturalHeight, natural);
            double height;
            double baseline;
            var rule = format.LineSpacingRule ?? LineSpacingRule.Auto;
            var spacing = format.LineSpacing ?? 1.0;

            // Word's leading placement, verified against its output: "exactly" and "at least"
            // align the text to the bottom of the line box (extra space above), while "multiple"
            // (auto) keeps the baseline at the font ascent and adds the extra space below.
            if (rule == LineSpacingRule.Exact && spacing > 0)
            {
                height = spacing;
                baseline = Math.Max(0, height - line.Descent);
            }
            else if (rule == LineSpacingRule.AtLeast && spacing > 0)
            {
                height = Math.Max(spacing, box);
                baseline = Math.Max(line.Ascent, height - line.Descent);
            }
            else
            {
                // The multiple scales the font-based line height only. An inline picture
                // contributes its box plus the leading the multiple adds to its run's font
                // height — (multiple − 1) × font height — never box times the multiple
                // (PROBE11/12: below a 240px image, line=276 adds ~2px, line=360 ~9px,
                // single spacing nothing; identical amid text; the extra hangs below).
                var multiple = spacing <= 0 ? 1.0 : spacing;
                // Sub-single AUTO multiples compress below the glyph height (sample3:
                // w:line=120 spacers render 6.7pt in Word) — keep in sync with
                // LineBoxHeight.
                height = multiple < 1
                    ? Math.Max(natural, line.FontHeight) * multiple
                    : Math.Max(natural, line.FontHeight * multiple);
                if (line.ImageBox > 0)
                    height = Math.Max(height,
                        line.ImageBox + Math.Max(0, multiple - 1) * line.ImageRunFontHeight);
                baseline = line.Ascent;
            }
            if (baseline > height)
                baseline = height;

            // Paragraph borders occupy vertical space: the border sits w:space points
            // away from the text and its stroke width adds to the paragraph height
            // (sample1 p1: the Title's rule is sz=8 space=4 — Word renders the body
            // 5pt lower than a space-less border would).
            var lineBorders = format.Borders;
            double topExtra = 0, bottomExtra = 0;
            if (isFirst && borderTop && lineBorders != null && lineBorders.Top != null && lineBorders.Top.IsVisible)
                topExtra = lineBorders.Top.SpacePt + lineBorders.Top.WidthPt;
            if (isLast && borderBottom && lineBorders != null && lineBorders.Bottom != null && lineBorders.Bottom.IsVisible)
                bottomExtra = lineBorders.Bottom.SpacePt + lineBorders.Bottom.WidthPt;
            baseline += topExtra;

            var fragment = new Fragment(height + topExtra + bottomExtra);

            // The leading an auto line-spacing multiple adds hangs BELOW the ink and
            // does not block a page-bottom fit: Word places a 42nd line on sample4's
            // pages where 42 × 14.49pt + spacing overruns the 648pt body by 0.6pt —
            // the text itself (ascent+descent 13.43) fits, only the 1.06pt of
            // multiple-leading crosses the margin. The same holds under an inline
            // picture (sample4 p71: empty line + full-page photo share the page in
            // Word; our +1.06 pushed the photo off, leaving a blank page). Exact and
            // AtLeast rules keep their full height.
            if ((format.LineSpacingRule ?? LineSpacingRule.Auto) == LineSpacingRule.Auto
                && bottomExtra <= 0)
            {
                var inkHeight = Math.Max(Math.Max(line.Ascent + line.Descent, line.NaturalHeight), line.ImageBox);
                fragment.BottomSlackPt = Math.Max(0, height - inkHeight);
            }

            // Paragraph shading spans the whole content width.
            if (format.Shading.HasValue)
                fragment.Add(new RectOp { X = 0, Y = 0, Width = availableWidth, Height = height + topExtra + bottomExtra, Color = format.Shading.Value });

            // Horizontal alignment.
            var alignment = format.Alignment ?? TextAlignment.Left;
            var free = line.MaxWidth - line.ContentWidth;
            if (free < 0)
                free = 0;
            var x = line.Left;
            var justifyExtra = 0.0;

            if (!line.HasTab)
            {
                if (alignment == TextAlignment.Center)
                {
                    x += free / 2;
                }
                else if (alignment == TextAlignment.Right)
                {
                    x += free;
                }
                else if (alignment == TextAlignment.Justify && !isLast && !line.EndsWithBreak)
                {
                    var spaces = 0;
                    for (var i = 0; i < line.Atoms.Count; i++)
                    {
                        if (line.Atoms[i].IsSpace && i < LastContentIndex(line))
                            spaces++;
                    }
                    if (spaces > 0 && free > 0 && free < line.MaxWidth * 0.5)
                        justifyExtra = free / spaces;
                }
            }

            var lastContent = LastContentIndex(line);
            for (var i = 0; i < line.Atoms.Count; i++)
            {
                var atom = line.Atoms[i];
                var width = atom.Width;
                if (justifyExtra > 0 && atom.IsSpace && i < lastContent)
                    width += justifyExtra;

                EmitAtom(fragment, atom, x, baseline, width, height);
                x += width;
            }

            EmitParagraphBorders(fragment, format, availableWidth, height + topExtra + bottomExtra,
                                 isFirst && borderTop, isLast && borderBottom);
            return fragment;
        }

        /// <summary>True when two paragraphs' border sets are identical (Word merges them).</summary>
        internal static bool SameBorders(Borders a, Borders b)
        {
            if (a == null || b == null)
                return false;
            return SameBorder(a.Top, b.Top) && SameBorder(a.Bottom, b.Bottom)
                   && SameBorder(a.Left, b.Left) && SameBorder(a.Right, b.Right);
        }

        private static bool SameBorder(Border a, Border b)
        {
            var aVisible = a != null && a.IsVisible;
            var bVisible = b != null && b.IsVisible;
            if (!aVisible || !bVisible)
                return aVisible == bVisible;
            return a.Style == b.Style && Math.Abs(a.WidthPt - b.WidthPt) < 0.01 && a.Color == b.Color;
        }

        private static int LastContentIndex(Line line)
        {
            for (var i = line.Atoms.Count - 1; i >= 0; i--)
            {
                if (!line.Atoms[i].IsSpace)
                    return i;
            }
            return -1;
        }

        private void EmitAtom(Fragment fragment, Atom atom, double x, double baseline, double width, double lineHeight)
        {
            var bookmark = atom as BookmarkAtom;
            if (bookmark != null)
            {
                fragment.Add(new BookmarkOp { X = x, Y = 0, Name = bookmark.Name });
                return;
            }

            var anchor = atom as AnchorAtom;
            if (anchor != null)
            {
                fragment.Add(new AnchoredOp
                {
                    X = x,
                    Y = 0,
                    Width = anchor.Source.WidthPt,
                    Height = anchor.FrameHeight,
                    Content = anchor.Content,
                    HorizontalFrom = anchor.Source.HorizontalFrom,
                    HorizontalOffset = anchor.Source.HorizontalOffsetPt,
                    HorizontalAlign = anchor.Source.HorizontalAlign,
                    VerticalFrom = anchor.Source.VerticalFrom,
                    VerticalOffset = anchor.Source.VerticalOffsetPt,
                    VerticalAlign = anchor.Source.VerticalAlign,
                    BehindDoc = anchor.Source.BehindDoc,
                });
                return;
            }

            var image = atom as ImageAtom;
            if (image != null)
            {
                var top = baseline - atom.Ascent;
                if (image.Image != null)
                {
                    // The picture is drawn at its own extent inside the reserved effect box.
                    // A framed picture (a:ln + drop shadow) additionally shows the shadow's
                    // boundary as a soft grey edge around the extent (measured against Word's
                    // export; drawing the photo inset scored worse than the full extent).
                    var inset = image.Source.FrameWidthPt > 0 ? image.Source.FrameWidthPt / 2 : 0;
                    inset = Math.Min(inset, Math.Min(image.DrawWidth, image.DrawHeight) / 4);
                    fragment.Add(new ImageOp
                    {
                        X = x + image.Source.EffectLeftPt,
                        Y = top + image.Source.EffectTopPt,
                        Width = image.DrawWidth,
                        Height = image.DrawHeight,
                        Key = image.Source.PartName,
                        Image = image.Image,
                        RotationDeg = image.Source.RotationDeg,
                        CropLeft = image.Source.CropLeft,
                        CropTop = image.Source.CropTop,
                        CropRight = image.Source.CropRight,
                        CropBottom = image.Source.CropBottom,
                    });
                    if (inset > 0)
                    {
                        // The drop shadow's inner boundary reads as a soft grey edge around
                        // the white frame in Word's export.
                        var fx = x + image.Source.EffectLeftPt - inset;
                        var fy = top + image.Source.EffectTopPt - inset;
                        var fw = image.DrawWidth + inset * 2;
                        var fh = image.DrawHeight + inset * 2;
                        const uint edge = 0x9A9A9A;
                        fragment.Add(new LineOp { X1 = fx, Y1 = fy, X2 = fx + fw, Y2 = fy, Color = edge, Width = 0.8 });
                        fragment.Add(new LineOp { X1 = fx, Y1 = fy + fh, X2 = fx + fw, Y2 = fy + fh, Color = edge, Width = 1.2 });
                        fragment.Add(new LineOp { X1 = fx, Y1 = fy, X2 = fx, Y2 = fy + fh, Color = edge, Width = 0.8 });
                        fragment.Add(new LineOp { X1 = fx + fw, Y1 = fy, X2 = fx + fw, Y2 = fy + fh, Color = edge, Width = 1.2 });
                    }
                }
                // Annotation rectangles a group shape draws over the picture (highlight
                // boxes on screenshots); coordinates are box-relative points.
                if (image.Source.Overlays != null)
                {
                    foreach (var overlay in image.Source.Overlays)
                    {
                        var ox = x + overlay.X;
                        var oy = top + overlay.Y;
                        if (overlay.FillColor.HasValue)
                        {
                            fragment.Add(new RectOp
                            {
                                X = ox, Y = oy, Width = overlay.Width, Height = overlay.Height,
                                Color = overlay.FillColor.Value,
                            });
                        }
                        if (overlay.OutlineColor.HasValue)
                        {
                            var w = overlay.OutlineWidthPt > 0 ? overlay.OutlineWidthPt : 1;
                            var c = overlay.OutlineColor.Value;
                            fragment.Add(new LineOp { X1 = ox, Y1 = oy, X2 = ox + overlay.Width, Y2 = oy, Width = w, Color = c });
                            fragment.Add(new LineOp { X1 = ox, Y1 = oy + overlay.Height, X2 = ox + overlay.Width, Y2 = oy + overlay.Height, Width = w, Color = c });
                            fragment.Add(new LineOp { X1 = ox, Y1 = oy, X2 = ox, Y2 = oy + overlay.Height, Width = w, Color = c });
                            fragment.Add(new LineOp { X1 = ox + overlay.Width, Y1 = oy, X2 = ox + overlay.Width, Y2 = oy + overlay.Height, Width = w, Color = c });
                        }
                    }
                }

                AddLink(fragment, atom, x, top, atom.Width, atom.Ascent);
                return;
            }

            var tab = atom as TabAtom;
            if (tab != null)
            {
                if (tab.Leader != TabLeader.None && width > 2)
                    EmitLeader(fragment, tab, x, baseline, width);
                return;
            }

            var text = atom as TextAtom;
            if (text == null || string.IsNullOrEmpty(text.Text))
                return;

            if (text.Highlight.HasValue)
            {
                fragment.Add(new RectOp
                {
                    X = x,
                    Y = baseline - atom.Ascent,
                    Width = width,
                    Height = atom.Ascent + atom.Descent,
                    Color = text.Highlight.Value,
                });
            }

            if (!text.IsSpace || text.Underline != UnderlineStyle.None)
            {
                fragment.Add(new TextOp
                {
                    X = x,
                    Y = baseline + text.BaselineShift,
                    Font = text.Font,
                    Size = text.Size,
                    Color = text.Color,
                    Text = text.Text,
                    CharSpacing = text.CharSpacing,
                    Field = text.Field,
                });
            }

            if (text.Underline != UnderlineStyle.None)
            {
                var thickness = Math.Max(0.4, text.Size * 0.055);
                var y = baseline + text.BaselineShift + text.Size * 0.12;
                fragment.Add(new LineOp
                {
                    X1 = x,
                    Y1 = y,
                    X2 = x + width,
                    Y2 = y,
                    Width = thickness,
                    Color = text.UnderlineColor,
                    Style = text.Underline == UnderlineStyle.Dotted ? BorderStyle.Dotted
                          : text.Underline == UnderlineStyle.Dashed ? BorderStyle.Dashed
                          : BorderStyle.Single,
                });
                if (text.Underline == UnderlineStyle.Double)
                {
                    fragment.Add(new LineOp
                    {
                        X1 = x, Y1 = y + thickness * 2, X2 = x + width, Y2 = y + thickness * 2,
                        Width = thickness, Color = text.UnderlineColor,
                    });
                }
            }

            if (text.Strike)
            {
                var y = baseline + text.BaselineShift - text.Size * 0.26;
                fragment.Add(new LineOp
                {
                    X1 = x, Y1 = y, X2 = x + width, Y2 = y,
                    Width = Math.Max(0.4, text.Size * 0.05), Color = text.Color,
                });
            }

            AddLink(fragment, atom, x, baseline - atom.Ascent, width, atom.Ascent + atom.Descent);
        }

        private void EmitLeader(Fragment fragment, TabAtom tab, double x, double baseline, double width)
        {
            var leaderChar = tab.Leader == TabLeader.Dot ? "." : (tab.Leader == TabLeader.Hyphen ? "-" : "_");
            var font = tab.Font;
            var unit = font.Measure(leaderChar, tab.Size);
            if (unit <= 0.1)
                return;

            var count = (int)Math.Floor((width - unit) / unit);
            if (count <= 0)
                return;
            if (count > 500)
                count = 500;

            var text = new string(leaderChar[0], count);
            var used = count * unit;
            fragment.Add(new TextOp
            {
                X = x + (width - used),
                Y = baseline,
                Font = font,
                Size = tab.Size,
                Color = tab.Color,
                Text = text,
            });
        }

        private void AddLink(Fragment fragment, Atom atom, double x, double y, double width, double height)
        {
            if (!_ctx.Options.CreateHyperlinks)
                return;
            if (string.IsNullOrEmpty(atom.LinkUrl) && string.IsNullOrEmpty(atom.LinkAnchor))
                return;
            if (width <= 0 || height <= 0)
                return;

            fragment.Add(new LinkOp
            {
                X = x,
                Y = y,
                Width = width,
                Height = height,
                Url = atom.LinkUrl,
                Anchor = atom.LinkAnchor,
            });
        }

        private static void EmitParagraphBorders(Fragment fragment, ParagraphFormat format, double width,
                                                 double height, bool isFirst, bool isLast)
        {
            var borders = format.Borders;
            if (borders == null)
                return;

            // The strokes sit just inside the space the paragraph reserved for them.
            if (isFirst && borders.Top != null && borders.Top.IsVisible)
                fragment.Add(HorizontalLine(0, width, borders.Top.WidthPt / 2, borders.Top));
            if (isLast && borders.Bottom != null && borders.Bottom.IsVisible)
                fragment.Add(HorizontalLine(0, width, height - borders.Bottom.WidthPt / 2, borders.Bottom));
            if (borders.Left != null && borders.Left.IsVisible)
                fragment.Add(new LineOp
                {
                    X1 = 0, Y1 = 0, X2 = 0, Y2 = height,
                    Width = borders.Left.WidthPt, Color = borders.Left.Color, Style = borders.Left.Style,
                });
            if (borders.Right != null && borders.Right.IsVisible)
                fragment.Add(new LineOp
                {
                    X1 = width, Y1 = 0, X2 = width, Y2 = height,
                    Width = borders.Right.WidthPt, Color = borders.Right.Color, Style = borders.Right.Style,
                });
        }

        private static LineOp HorizontalLine(double x1, double x2, double y, Border border)
        {
            return new LineOp
            {
                X1 = x1, Y1 = y, X2 = x2, Y2 = y,
                Width = border.WidthPt, Color = border.Color, Style = border.Style,
            };
        }
    }
}
