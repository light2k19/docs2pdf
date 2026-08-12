using System;
using System.Collections.Generic;
using System.Globalization;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    /// <summary>Places fragments onto pages, adds headers and footers, and collects outline data.</summary>
    internal sealed class DocumentLayout
    {
        private readonly LayoutContext _ctx;
        private readonly LayoutResult _result = new LayoutResult();
        private readonly List<SectionProperties> _pageSections = new List<SectionProperties>();
        private readonly List<bool> _pageIsSectionStart = new List<bool>();

        private LayoutPage _page;
        private SectionProperties _props;
        private double _y;
        /// <summary>The kept top spacing came from a break RUN (not w:pageBreakBefore) —
        /// legacy documents dissolve style-derived spacing there.</summary>
        private bool _topSpacingFromBreakRun;
        /// <summary>Leading spacing is kept at the top of this page (explicit break / section start).</summary>
        private bool _keepTopSpacing;
        private int _pageNumber = 1;
        private object _currentTableId;
        private readonly List<Fragment> _headerRows = new List<Fragment>();

        /// <summary>A fragment already placed on the current page, recorded so it can be pulled back.</summary>
        private struct PlacedFragment
        {
            public Fragment Fragment;
            public int OpStart;
            public double Y;
            public int OutlineCount;
            public string Bookmark;
        }

        private readonly List<PlacedFragment> _placed = new List<PlacedFragment>();

        /// <summary>Top of the body area on the current page; grows when the header is taller than the margin.</summary>
        private double _contentTop;
        /// <summary>Height available to body content on the current page.</summary>
        private double _contentHeight;
        /// <summary>Content height before any footnote reservation on the current page.</summary>
        private double _contentHeightFull;
        /// <summary>Footnote bodies to render at the bottom of the current page.</summary>
        private readonly List<Fragment> _pageNotes = new List<Fragment>();
        private readonly Dictionary<HeaderFooter, double> _headerFooterHeights = new Dictionary<HeaderFooter, double>();
        private LayoutBuilder _measureBuilder;

        public DocumentLayout(LayoutContext context)
        {
            _ctx = context;
        }

        public LayoutResult Layout(WordDocument document)
        {
            _ctx.DefaultTabStopPt = document.DefaultTabStopPt;
            _ctx.LegacyCellSpacing = document.CompatibilityMode < 15;
            _ctx.DefaultParagraphStyleId = document.DefaultParagraphStyleId;
            var builder = new LayoutBuilder(_ctx);

            for (var sectionIndex = 0; sectionIndex < document.Sections.Count; sectionIndex++)
            {
                var section = document.Sections[sectionIndex];
                _props = section.Properties ?? new SectionProperties();

                if (_props.PageNumberStart.HasValue)
                    _pageNumber = _props.PageNumberStart.Value;

                builder.MaxBlockHeight = Math.Max(48, SmallestContentHeight(_props));
                // A multi-column section flows its blocks at column width and composes
                // them into one balanced side-by-side block (sample3: the continuous
                // 2-column tail sits at the bottom of Word's last page).
                var fragments = _props.Columns > 1
                    ? BuildColumnBlock(builder, section, _props)
                    : builder.BuildBlocks(section.Blocks, Math.Max(1, _props.ContentWidthPt));

                // A continuous section carries on down the same page; only explicit
                // (nextPage-style) section breaks start a fresh one.
                if (_page == null || (sectionIndex > 0 && section.StartsNewPage))
                    NewPage(true);

                PlaceAll(fragments);
            }

            if (document.Notes.Count > 0)
                PlaceNotes(document, builder);

            if (_page == null)
                NewPage(true);
            FlushPageNotes();

            RenderHeadersAndFooters();
            SubstituteFields();
            return _result;
        }

        /// <summary>
        /// Flows a multi-column section's blocks at column width and composes them into
        /// one fragment with the columns side by side, balanced the way Word closes a
        /// continuous section: fill each column to roughly total/n before moving on.
        /// </summary>
        private List<Fragment> BuildColumnBlock(LayoutBuilder builder, Section section, SectionProperties props)
        {
            var n = props.Columns;
            var spacing = props.ColumnSpacingPt;
            var colWidth = Math.Max(1, (props.ContentWidthPt - (n - 1) * spacing) / n);
            var parts = builder.BuildBlocks(section.Blocks, colWidth);

            double total = 0;
            foreach (var part in parts)
                total += part.Height;
            var target = total / n;

            var block = new Fragment(0);
            var column = 0;
            double y = 0, maxBottom = 0;
            foreach (var part in parts)
            {
                if (column < n - 1 && y > 0 && y + part.Height > target + 0.5)
                {
                    column++;
                    y = 0;
                }
                // Spacing dissolves at a column top like at a page top.
                if (part.IsSpacing && y <= 0)
                    continue;
                var dx = column * (colWidth + spacing);
                foreach (var op in part.Ops)
                    block.Ops.Add(op.Translate(dx, y));
                y += part.Height;
                if (y > maxBottom)
                    maxBottom = y;
            }
            block.Height = maxBottom;
            return new List<Fragment> { block };
        }

        private void PlaceNotes(WordDocument document, LayoutBuilder builder)
        {
            var blocks = new List<Block>();
            foreach (var note in document.Notes)
                blocks.AddRange(note.Value);
            if (blocks.Count == 0)
                return;

            var fragments = builder.BuildBlocks(blocks, Math.Max(1, _props.ContentWidthPt));

            // Separator rule above the collected notes.
            var separator = new Fragment(10);
            separator.Add(new LineOp
            {
                X1 = 0, Y1 = 6, X2 = _props.ContentWidthPt / 3, Y2 = 6, Width = 0.5, Color = 0x808080,
            });
            fragments.Insert(0, separator);
            PlaceAll(fragments);
        }

        private void PlaceAll(List<Fragment> fragments)
        {
            // Trailing empty paragraphs must not spill onto a page of their own.
            var blankFromHere = new bool[fragments.Count + 1];
            blankFromHere[fragments.Count] = true;
            for (var i = fragments.Count - 1; i >= 0; i--)
                blankFromHere[i] = blankFromHere[i + 1] && IsBlank(fragments[i]);

            var breakPending = false;
            object breakSourceId = null;
            var debugPlace = Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_PLACE") != null;
            for (var i = 0; i < fragments.Count; i++)
            {
                var fragment = fragments[i];
                if (debugPlace)
                    Console.Error.WriteLine("place pg={0} y={1:F1} h={2:F1} slack={3:F1} cH={4:F1} spc={5}{6} brB={7} keepN={8} row={9} ops={10}",
                        _result.Pages.Count, _y, fragment.Height, fragment.BottomSlackPt, _contentHeight,
                        fragment.IsSpacing ? 1 : 0, fragment.IsSpaceBefore ? "b" : "",
                        fragment.PageBreakBefore ? 1 : 0, fragment.KeepWithNext ? 1 : 0,
                        fragment.RowSource != null ? 1 : 0, fragment.Ops.Count);

                // The repeated table header must not leak onto a page the table never
                // reaches: clear the table state as soon as the flow moves past the table,
                // BEFORE any page break this fragment causes (p15: a heading pushed to a
                // fresh page used to drag the Features header row along with it).
                if (fragment.TableId == null && _currentTableId != null)
                {
                    _currentTableId = null;
                    _headerRows.Clear();
                }

                // Break only when something already sits on this page; page borders and
                // headers do not count, so a break at the top of a page stays put.
                // A break with nothing but whitespace after it is dropped — Word collapses a
                // trailing page break with the section break that follows it. And when the
                // break-carrying paragraph is itself the only thing on the page (it spilled
                // over), the page break has already happened and is not applied again.
                if ((fragment.PageBreakBefore || breakPending) && !blankFromHere[i]
                    && (_y > 0 || _placed.Count > 0)
                    && !OnlyParagraphOnPage(breakPending ? breakSourceId : fragment.ParagraphId))
                {
                    NewPage(false);
                    _keepTopSpacing = true;  // an explicit break keeps the following space-before
                    // ...except style-derived spacing after a break RUN in legacy docs
                    // (see the dissolve below). A w:pageBreakBefore paragraph always
                    // keeps its spacing (sample1's Heading1 chapter tops).
                    _topSpacingFromBreakRun = breakPending && !fragment.PageBreakBefore;
                }
                breakPending = false;
                breakSourceId = null;

                // Spilled-over space-after is dropped: at the top of a page, and also when it
                // does not fit at the bottom (it must never cause the page break itself and
                // reappear as leading space on the next page).
                if (fragment.IsSpacing && !fragment.IsSpaceBefore
                    && (_y <= 0 || fragment.Height > _contentHeight - _y + 0.5))
                    continue;

                // Space-before is suppressed at the top of a page reached by natural flow;
                // it survives after an explicit page break or at a section start (PROBE5:
                // Word starts the carried paragraph at the content top, license doc keeps
                // spacing after hard breaks). LEGACY documents dissolve STYLE-derived
                // spacing after a break RUN (sample3 p2: the pie paragraph starts at the
                // content top although its Normal style carries before=6pt and p1 ends
                // with a <w:br type="page"/>) — but a w:pageBreakBefore paragraph keeps
                // its style spacing (sample1's Heading1 pages sit 24pt down like Word).
                if (fragment.IsSpacing && fragment.IsSpaceBefore && _y <= 0
                    && (!_keepTopSpacing
                        || (_topSpacingFromBreakRun && _ctx.LegacyCellSpacing && !fragment.SpacingIsDirect)))
                    continue;

                var available = _contentHeight - _y;
                var needed = fragment.Height;

                // Keep-with-next: the paragraph and the first lines that follow it must share a page.
                if (fragment.KeepWithNext && needed <= _contentHeight)
                {
                    var chain = needed;
                    var linesAfter = 0;
                    for (var j = i + 1; j < fragments.Count && j <= i + 8; j++)
                    {
                        chain += fragments[j].Height;
                        if (fragments[j].IsSpacing)
                            continue;
                        linesAfter++;
                        // Two lines of the following paragraph, matching Word's widow control.
                        if (linesAfter >= 2 || !fragments[j].KeepWithNext && linesAfter >= 1 && !fragments[j].WidowControl)
                            break;
                    }
                    if (chain > available && chain <= _contentHeight && _y > 0)
                    {
                        NewPage(false);
                        // The chain reaches the new page by natural flow, so a leading
                        // space-before dissolves at the page top exactly as in direct flow
                        // (MyLesen p6: the keepNext-moved section-2 heading starts at the
                        // content top in Word, not 12pt down).
                        if (fragment.IsSpacing && fragment.IsSpaceBefore && !_keepTopSpacing)
                            continue;
                    }
                }

                // A tall table row is split across pages rather than pushed onto the next one.
                if (fragment.RowSource != null && fragment.Height > _contentHeight - _y + 0.5)
                {
                    PlaceSplitRow(fragment);
                    breakPending = fragment.PageBreakAfter;
                    if (breakPending)
                        breakSourceId = fragment.ParagraphId;
                    continue;
                }

                // The fit tolerance is the LARGER of the pixel epsilon and the line's
                // below-text leading — never their sum (sample4: summing them accepted
                // a 42nd line on p4 that Word breaks; the text box itself must fit).
                if (fragment.Height - _contentHeight + _y > Math.Max(0.5, fragment.BottomSlackPt) && _y > 0)
                {
                    if (blankFromHere[i])
                        return;              // nothing but whitespace is left

                    // Widow/orphan control: never strand a single line of a paragraph on either page.
                    var pullBack = WidowOrphanPullBack(fragment);
                    // A line already at the top of its page is not stranded — pulling it
                    // would only leave a blank page behind (sample4: a full-page photo
                    // that opens its page stays put; the next photo-line starts fresh).
                    if (pullBack >= _placed.Count)
                        pullBack = 0;
                    var moved = PullBack(pullBack);
                    NewPage(false);
                    foreach (var earlier in moved)
                    {
                        // Pulled-back spacing dissolves at the top of the new page the same
                        // way it would had the paragraph flowed there directly.
                        if (earlier.IsSpacing && _y <= 0 && !_keepTopSpacing)
                            continue;
                        Place(earlier);
                    }
                    // The spilled fragment itself dissolves too when it is bare space-before
                    // landing at the top of the new page (natural flow).
                    if (fragment.IsSpacing && fragment.IsSpaceBefore && _y <= 0 && !_keepTopSpacing)
                        continue;

                    // The reunited chain can still overflow when the following line is
                    // taller than what the fresh page has left (sample4: an orphan photo
                    // pulled to its own page, chased by a full-page photo) — the fragment
                    // then starts yet another page and the pulled lines stay behind,
                    // exactly like Word's p89 (photo alone) / p90 (full-page photo).
                    if (fragment.Height - _contentHeight + _y > Math.Max(0.5, fragment.BottomSlackPt) && _y > 0)
                        NewPage(false);
                }

                Place(fragment);
                breakPending = fragment.PageBreakAfter;
                if (breakPending)
                    breakSourceId = fragment.ParagraphId;
            }
        }

        /// <summary>True when every fragment on the current page belongs to the given paragraph.</summary>
        private bool OnlyParagraphOnPage(object paragraphId)
        {
            if (paragraphId == null || _placed.Count == 0)
                return false;
            foreach (var placed in _placed)
            {
                if (!ReferenceEquals(placed.Fragment.ParagraphId, paragraphId))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// How many already-placed fragments must move to the next page with this one:
        /// one line left behind at the bottom (orphan) or a lone last line moving up (widow).
        /// </summary>
        private int WidowOrphanPullBack(Fragment fragment)
        {
            if (!fragment.WidowControl || fragment.ParagraphId == null || fragment.LineCount < 2 || fragment.IsSpacing)
                return 0;

            int placedLines;
            var placedFragments = TrailingFragmentsOf(fragment.ParagraphId, out placedLines);
            if (placedLines == 0)
                return 0;

            // Orphan: only the first line would remain on this page.
            if (placedLines == 1 && fragment.LineIndex == 1)
                return placedFragments;

            // Widow: the last line would sit alone on the next page.
            if (fragment.LineIndex == fragment.LineCount - 1 && placedLines >= 2)
            {
                var take = 0;
                var lines = 0;
                for (var i = _placed.Count - 1; i >= 0 && lines < 1; i--)
                {
                    if (!ReferenceEquals(_placed[i].Fragment.ParagraphId, fragment.ParagraphId))
                        break;
                    take++;
                    if (!_placed[i].Fragment.IsSpacing)
                        lines++;
                }
                return take;
            }

            return 0;
        }

        /// <summary>Places a table row that does not fit, slicing it over as many pages as needed.</summary>
        private void PlaceSplitRow(Fragment fragment)
        {
            var source = fragment.RowSource;

            // Rows Word would keep together simply move to the next page.
            if (source.CantSplit)
            {
                if (_y > 0)
                    NewPage(false);
                Place(fragment);
                return;
            }

            // Otherwise the row starts here if even one line fits, exactly like Word.
            source.Rewind();
            var guard = 0;
            while (!source.Exhausted && guard++ < 4000)
            {
                var available = _contentHeight - _y;
                var part = available > 1 ? source.Next(available, _y <= 0) : null;
                if (part == null)
                {
                    if (_y <= 0)
                        break;               // nothing fits even on an empty page
                    NewPage(false);
                    continue;
                }
                part.RowSource = null;
                Place(part);
                // A cut always continues on the next page: the slice hugs its content, so
                // room may remain below it, but Word never resumes a split row on the same
                // page (the widow/orphan pull-back would otherwise be re-taken right here).
                if (!source.Exhausted)
                    NewPage(false);
            }
        }

        /// <summary>True when a fragment draws nothing (an empty paragraph or bare spacing).</summary>
        private static bool IsBlank(Fragment fragment)
        {
            foreach (var op in fragment.Ops)
            {
                if (!(op is BookmarkOp))
                    return false;
            }
            return true;
        }

        private void Place(Fragment fragment)
        {
            if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_PAGE") ==
                _result.Pages.Count.ToString(System.Globalization.CultureInfo.InvariantCulture))
            {
                string text = null;
                foreach (var op in fragment.Ops)
                {
                    var t = op as TextOp;
                    if (t != null) { text = t.Text; break; }
                }
                Console.Error.WriteLine("place p{0} y={1:F1} h={2:F1} ops={3} spacing={4} row={5} anchor={6} text='{7}'",
                    _result.Pages.Count, _y, fragment.Height, fragment.Ops.Count, fragment.IsSpacing,
                    fragment.RowSource != null,
                    fragment.Ops.Count > 0 && fragment.Ops[0] is AnchoredOp,
                    text ?? string.Empty);
            }
            if (fragment.TableId != null)
            {
                if (!ReferenceEquals(fragment.TableId, _currentTableId))
                {
                    _currentTableId = fragment.TableId;
                    _headerRows.Clear();
                }
                if (fragment.IsTableHeaderRow && _headerRows.Count < 4 && _y > 0)
                    _headerRows.Add(fragment);
            }
            else if (_currentTableId != null)
            {
                _currentTableId = null;
                _headerRows.Clear();
            }

            var dx = _props.MarginLeftPt;
            var dy = _contentTop + _y;

            var record = new PlacedFragment
            {
                Fragment = fragment,
                OpStart = _page.Ops.Count,
                Y = _y,
                OutlineCount = _result.Outline.Count,
            };

            foreach (var op in fragment.Ops)
            {
                var bookmark = op as BookmarkOp;
                if (bookmark != null)
                {
                    if (!string.IsNullOrEmpty(bookmark.Name) && !_result.Bookmarks.ContainsKey(bookmark.Name))
                    {
                        _result.Bookmarks[bookmark.Name] =
                            new KeyValuePair<int, double>(_result.Pages.Count - 1, dy + bookmark.Y);
                        record.Bookmark = record.Bookmark ?? bookmark.Name;
                    }
                    continue;
                }

                var anchored = op as AnchoredOp;
                if (anchored != null)
                {
                    PlaceAnchored((AnchoredOp)anchored.Translate(dx, dy));
                    continue;
                }

                _page.Ops.Add(op.Translate(dx, dy));
            }

            if (fragment.HeadingLevel > 0 && !string.IsNullOrEmpty(fragment.HeadingText))
            {
                _result.Outline.Add(new OutlineEntry
                {
                    Level = fragment.HeadingLevel,
                    Title = fragment.HeadingText,
                    PageIndex = _result.Pages.Count - 1,
                    Y = dy,
                });
            }

            _y += fragment.Height;
            _placed.Add(record);

            // Footnotes referenced by this fragment claim space at the page bottom:
            // later fragments on this page see the reduced content height, and the
            // notes render under a separator when the page finishes.
            if (fragment.FootnoteBlocks != null)
            {
                foreach (var note in fragment.FootnoteBlocks)
                {
                    if (_pageNotes.Count == 0)
                        _contentHeight -= NoteSeparatorHeight;
                    _pageNotes.Add(note);
                    _contentHeight -= note.Height;
                }
                fragment.FootnoteBlocks = null;
            }
        }

        private const double NoteSeparatorHeight = 10;

        /// <summary>Renders the current page's footnotes at its bottom, above the footer.</summary>
        private void FlushPageNotes()
        {
            if (_page == null || _pageNotes.Count == 0)
                return;

            var height = NoteSeparatorHeight;
            foreach (var note in _pageNotes)
                height += note.Height;

            var dx = _props.MarginLeftPt;
            var y = _contentTop + _contentHeightFull - height;
            _page.Ops.Add(new LineOp
            {
                X1 = dx, Y1 = y + 6, X2 = dx + _props.ContentWidthPt / 3, Y2 = y + 6,
                Width = 0.5, Color = 0x808080,
            });
            y += NoteSeparatorHeight;
            foreach (var note in _pageNotes)
            {
                foreach (var op in note.Ops)
                    _page.Ops.Add(op.Translate(dx, y));
                y += note.Height;
            }
            _pageNotes.Clear();
        }

        /// <summary>Resolves a floating frame's anchor into page coordinates and draws it.</summary>
        private void PlaceAnchored(AnchoredOp op)
        {
            var width = op.Width > 0 ? op.Width : 0;
            var height = op.Height > 0 ? op.Height : 0;

            double origin, extent;
            HorizontalReference(op.HorizontalFrom, op.X, out origin, out extent);
            var x = string.IsNullOrEmpty(op.HorizontalAlign)
                ? origin + op.HorizontalOffset
                : AlignWithin(origin, extent, width, op.HorizontalAlign);

            VerticalReference(op.VerticalFrom, op.Y, out origin, out extent);
            var y = string.IsNullOrEmpty(op.VerticalAlign)
                ? origin + op.VerticalOffset
                : AlignWithin(origin, extent, height, op.VerticalAlign);

            // Keep the frame on the page rather than letting it disappear off an edge.
            x = Math.Max(-width, Math.Min(x, _props.PageWidthPt));
            y = Math.Max(-height, Math.Min(y, _props.PageHeightPt));

            var index = op.BehindDoc ? 0 : _page.Ops.Count;
            foreach (var content in op.Content)
            {
                var translated = content.Translate(x, y);
                if (op.BehindDoc)
                {
                    _page.Ops.Insert(index++, translated);
                }
                else
                {
                    // Word stacks floating frames above the body, including inline images
                    // placed later on the page; the renderer paints overlay ops last.
                    translated.Overlay = true;
                    _page.Ops.Add(translated);
                }
            }
        }

        private void HorizontalReference(AnchorRelativeFrom from, double anchorX, out double origin, out double extent)
        {
            switch (from)
            {
                case AnchorRelativeFrom.Page:
                    origin = 0;
                    extent = _props.PageWidthPt;
                    break;
                case AnchorRelativeFrom.LeftMargin:
                case AnchorRelativeFrom.InsideMargin:
                    origin = 0;
                    extent = _props.MarginLeftPt;
                    break;
                case AnchorRelativeFrom.RightMargin:
                case AnchorRelativeFrom.OutsideMargin:
                    origin = _props.PageWidthPt - _props.MarginRightPt;
                    extent = _props.MarginRightPt;
                    break;
                case AnchorRelativeFrom.Character:
                    origin = anchorX;
                    extent = _props.PageWidthPt - _props.MarginRightPt - anchorX;
                    break;
                default:                       // margin, column
                    origin = _props.MarginLeftPt;
                    extent = _props.ContentWidthPt;
                    break;
            }
        }

        private void VerticalReference(AnchorRelativeFrom from, double anchorY, out double origin, out double extent)
        {
            switch (from)
            {
                case AnchorRelativeFrom.Page:
                    origin = 0;
                    extent = _props.PageHeightPt;
                    break;
                case AnchorRelativeFrom.TopMargin:
                    origin = 0;
                    extent = _contentTop;
                    break;
                case AnchorRelativeFrom.BottomMargin:
                    origin = _props.PageHeightPt - _props.MarginBottomPt;
                    extent = _props.MarginBottomPt;
                    break;
                case AnchorRelativeFrom.Margin:
                    origin = _contentTop;
                    extent = _contentHeight;
                    break;
                default:                       // paragraph, line, text
                    origin = anchorY;
                    extent = _contentHeight;
                    break;
            }
        }

        private static double AlignWithin(double origin, double extent, double size, string align)
        {
            switch (align)
            {
                case "center":
                case "centre":
                    return origin + (extent - size) / 2;
                case "right":
                case "bottom":
                case "outside":
                    return origin + extent - size;
                default:                       // left, top, inside
                    return origin;
            }
        }

        /// <summary>
        /// Removes the last <paramref name="count"/> fragments from the current page
        /// and returns them, so they can be re-placed on the next page (widow/orphan control).
        /// </summary>
        private List<Fragment> PullBack(int count)
        {
            var moved = new List<Fragment>();
            if (count <= 0 || _placed.Count == 0)
                return moved;

            count = Math.Min(count, _placed.Count);
            var first = _placed[_placed.Count - count];

            for (var i = _placed.Count - count; i < _placed.Count; i++)
            {
                moved.Add(_placed[i].Fragment);
                if (_placed[i].Bookmark != null)
                    _result.Bookmarks.Remove(_placed[i].Bookmark);
            }

            if (first.OpStart < _page.Ops.Count)
                _page.Ops.RemoveRange(first.OpStart, _page.Ops.Count - first.OpStart);
            if (first.OutlineCount < _result.Outline.Count)
                _result.Outline.RemoveRange(first.OutlineCount, _result.Outline.Count - first.OutlineCount);

            _y = first.Y;
            _placed.RemoveRange(_placed.Count - count, count);
            return moved;
        }

        /// <summary>Number of trailing fragments on this page that belong to the given paragraph.</summary>
        private int TrailingFragmentsOf(object paragraphId, out int lineCount)
        {
            var count = 0;
            lineCount = 0;
            for (var i = _placed.Count - 1; i >= 0; i--)
            {
                if (!ReferenceEquals(_placed[i].Fragment.ParagraphId, paragraphId))
                    break;
                count++;
                if (!_placed[i].Fragment.IsSpacing)
                    lineCount++;
            }
            return count;
        }

        private void NewPage(bool sectionStart)
        {
            FlushPageNotes();
            if (_ctx.Options.MaxPages > 0 && _result.Pages.Count >= _ctx.Options.MaxPages)
            {
                _ctx.Warn("The page limit of " + _ctx.Options.MaxPages.ToString(CultureInfo.InvariantCulture)
                          + " was reached; the remaining content was dropped.");
                _y = 0;
                return;
            }

            _page = new LayoutPage
            {
                Width = _props.PageWidthPt,
                Height = _props.PageHeightPt,
                Number = _pageNumber,
            };
            _pageNumber++;
            _result.Pages.Add(_page);
            _pageSections.Add(_props);
            _pageIsSectionStart.Add(sectionStart);
            _y = 0;
            _placed.Clear();
            // Natural flow breaks suppress leading spacing on the new page; explicit page
            // breaks and section starts keep it (the caller marks those).
            _keepTopSpacing = sectionStart;
            _topSpacingFromBreakRun = false;

            // Word grows the body area when the header or footer does not fit inside the margin.
            var headerHeight = MeasureHeaderFooter(SelectHeaderFooter(_props, sectionStart, _page.Number, true), _props, true);
            var footerHeight = MeasureHeaderFooter(SelectHeaderFooter(_props, sectionStart, _page.Number, false), _props, false);
            _contentTop = _props.MarginTopPt;
            if (headerHeight > 0)
                _contentTop = Math.Max(_contentTop, _props.HeaderDistancePt + headerHeight);
            var contentBottom = _props.MarginBottomPt;
            if (footerHeight > 0)
                contentBottom = Math.Max(contentBottom, _props.FooterDistancePt + footerHeight);
            _contentHeight = Math.Max(24, _props.PageHeightPt - _contentTop - contentBottom);
            _contentHeightFull = _contentHeight;

            DrawPageBorders(sectionStart);

            // Repeat table header rows when a table continues onto this page.
            if (_headerRows.Count > 0 && _currentTableId != null)
            {
                foreach (var header in _headerRows)
                {
                    foreach (var op in header.Ops)
                        _page.Ops.Add(op.Translate(_props.MarginLeftPt, _contentTop + _y));
                    _y += header.Height;
                }
            }
        }

        /// <summary>Height of a header or footer, measured once per part.</summary>
        private double MeasureHeaderFooter(HeaderFooter part, SectionProperties props, bool isHeader)
        {
            if (part == null || !_ctx.Options.RenderHeadersAndFooters)
                return 0;

            double height;
            if (_headerFooterHeights.TryGetValue(part, out height))
                return height;

            if (_measureBuilder == null)
                _measureBuilder = new LayoutBuilder(_ctx);
            _measureBuilder.MaxBlockHeight = Math.Max(24, props.PageHeightPt / 2);
            foreach (var fragment in _measureBuilder.BuildBlocks(part.Blocks, Math.Max(1, props.ContentWidthPt)))
            {
                // Row heights include their border widths. Word counts those in the space a
                // HEADER occupies (PROBE13, COM: body top 87.5pt = header content + 1pt of
                // borders), but the FOOTER anchors at the page bottom and its borders overlap
                // the body area — Word still fits a line there (SaBC p26: the embedded-map
                // image lands 0.45pt above the border-free footer limit).
                height += fragment.Height - (isHeader ? 0 : fragment.EdgeExtentPt);
                if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_HF") != null)
                    Console.Error.WriteLine("hf frag h={0:F2} edge={1:F2} spacing={2} row={3}",
                        fragment.Height, fragment.EdgeExtentPt, fragment.IsSpacing, fragment.RowSource != null);
            }

            _headerFooterHeights[part] = height;
            return height;
        }

        /// <summary>Worst-case body height for the section, used to bound splittable blocks.</summary>
        private double SmallestContentHeight(SectionProperties props)
        {
            var header = Math.Max(MeasureHeaderFooter(props.HeaderDefault, props, true),
                         Math.Max(MeasureHeaderFooter(props.HeaderFirst, props, true),
                                  MeasureHeaderFooter(props.HeaderEven, props, true)));
            var footer = Math.Max(MeasureHeaderFooter(props.FooterDefault, props, false),
                         Math.Max(MeasureHeaderFooter(props.FooterFirst, props, false),
                                  MeasureHeaderFooter(props.FooterEven, props, false)));

            var top = Math.Max(props.MarginTopPt, header > 0 ? props.HeaderDistancePt + header : 0);
            var bottom = Math.Max(props.MarginBottomPt, footer > 0 ? props.FooterDistancePt + footer : 0);
            return Math.Max(24, props.PageHeightPt - top - bottom);
        }

        /// <summary>Draws the section's page borders (w:pgBorders), honouring w:display.</summary>
        private void DrawPageBorders(bool isSectionFirstPage)
        {
            var borders = _props.PageBorders;
            if (borders == null || !borders.Any)
                return;

            var display = _props.PageBordersDisplay;
            if (string.Equals(display, "firstPage", StringComparison.OrdinalIgnoreCase) && !isSectionFirstPage)
                return;
            if (string.Equals(display, "notFirstPage", StringComparison.OrdinalIgnoreCase) && isSectionFirstPage)
                return;

            Func<Border, double, double> distance = (border, fallback) =>
                border == null ? fallback : border.SpacePt;

            double left, right, top, bottom;
            if (_props.PageBordersFromText)
            {
                left = _props.MarginLeftPt - distance(borders.Left, 0);
                right = _props.PageWidthPt - _props.MarginRightPt + distance(borders.Right, 0);
                top = _props.MarginTopPt - distance(borders.Top, 0);
                bottom = _props.PageHeightPt - _props.MarginBottomPt + distance(borders.Bottom, 0);
            }
            else
            {
                left = distance(borders.Left, 24);
                right = _props.PageWidthPt - distance(borders.Right, 24);
                top = distance(borders.Top, 24);
                bottom = _props.PageHeightPt - distance(borders.Bottom, 24);
            }

            if (borders.Top != null && borders.Top.IsVisible)
                _page.Ops.Add(new LineOp { X1 = left, Y1 = top, X2 = right, Y2 = top, Width = borders.Top.WidthPt, Color = borders.Top.Color, Style = borders.Top.Style });
            if (borders.Bottom != null && borders.Bottom.IsVisible)
                _page.Ops.Add(new LineOp { X1 = left, Y1 = bottom, X2 = right, Y2 = bottom, Width = borders.Bottom.WidthPt, Color = borders.Bottom.Color, Style = borders.Bottom.Style });
            if (borders.Left != null && borders.Left.IsVisible)
                _page.Ops.Add(new LineOp { X1 = left, Y1 = top, X2 = left, Y2 = bottom, Width = borders.Left.WidthPt, Color = borders.Left.Color, Style = borders.Left.Style });
            if (borders.Right != null && borders.Right.IsVisible)
                _page.Ops.Add(new LineOp { X1 = right, Y1 = top, X2 = right, Y2 = bottom, Width = borders.Right.WidthPt, Color = borders.Right.Color, Style = borders.Right.Style });
        }

        // ----------------------------------------------------------------- headers and footers

        private void RenderHeadersAndFooters()
        {
            if (!_ctx.Options.RenderHeadersAndFooters)
                return;

            var builder = new LayoutBuilder(_ctx);
            var total = _result.Pages.Count;

            for (var i = 0; i < _result.Pages.Count; i++)
            {
                var page = _result.Pages[i];
                var props = i < _pageSections.Count ? _pageSections[i] : _props;
                var isSectionStart = i < _pageIsSectionStart.Count && _pageIsSectionStart[i];

                var header = SelectHeaderFooter(props, isSectionStart, page.Number, true);
                var footer = SelectHeaderFooter(props, isSectionStart, page.Number, false);

                // Headers and footers are painted underneath the body, in their own drawing order.
                var underlay = new List<DrawOp>();

                if (header != null)
                {
                    builder.MaxBlockHeight = Math.Max(24, props.PageHeightPt / 2);
                    var fragments = builder.BuildBlocks(header.Blocks, Math.Max(1, props.ContentWidthPt));
                    var y = props.HeaderDistancePt;
                    foreach (var fragment in fragments)
                    {
                        foreach (var op in fragment.Ops)
                        {
                            if (op is BookmarkOp)
                                continue;
                            underlay.Add(op.Translate(props.MarginLeftPt, y));
                        }
                        y += fragment.Height;
                    }
                }

                if (footer != null)
                {
                    builder.MaxBlockHeight = Math.Max(24, props.PageHeightPt / 2);
                    var fragments = builder.BuildBlocks(footer.Blocks, Math.Max(1, props.ContentWidthPt));
                    double height = 0;
                    foreach (var fragment in fragments)
                        height += fragment.Height;

                    var y = props.PageHeightPt - props.FooterDistancePt - height;
                    foreach (var fragment in fragments)
                    {
                        foreach (var op in fragment.Ops)
                        {
                            if (op is BookmarkOp)
                                continue;
                            underlay.Add(op.Translate(props.MarginLeftPt, y));
                        }
                        y += fragment.Height;
                    }
                }

                if (underlay.Count > 0)
                    page.Ops.InsertRange(0, underlay);
            }

            _totalPages = total;
        }

        private int _totalPages;

        private static HeaderFooter SelectHeaderFooter(SectionProperties props, bool isSectionStart, int pageNumber, bool header)
        {
            if (props.TitlePage && isSectionStart)
            {
                var first = header ? props.HeaderFirst : props.FooterFirst;
                if (first != null)
                    return first;
                return null;
            }

            if (pageNumber % 2 == 0)
            {
                var even = header ? props.HeaderEven : props.FooterEven;
                if (even != null)
                    return even;
            }
            return header ? props.HeaderDefault : props.FooterDefault;
        }

        /// <summary>Applies the section's page number format (decimal, roman, letters).</summary>
        private static string FormatPageNumber(int number, SectionProperties props)
        {
            var format = props == null ? null : props.PageNumberFormat;
            if (string.IsNullOrEmpty(format))
                return number.ToString(CultureInfo.InvariantCulture);

            switch (format.ToLowerInvariant())
            {
                case "lowerroman": return Ooxml.NumberingTable.FormatNumber(number, Ooxml.NumberFormat.LowerRoman);
                case "upperroman": return Ooxml.NumberingTable.FormatNumber(number, Ooxml.NumberFormat.UpperRoman);
                case "lowerletter": return Ooxml.NumberingTable.FormatNumber(number, Ooxml.NumberFormat.LowerLetter);
                case "upperletter": return Ooxml.NumberingTable.FormatNumber(number, Ooxml.NumberFormat.UpperLetter);
                case "decimalzero": return Ooxml.NumberingTable.FormatNumber(number, Ooxml.NumberFormat.DecimalZero);
                default: return number.ToString(CultureInfo.InvariantCulture);
            }
        }

        /// <summary>Replaces PAGE / NUMPAGES placeholders with their final values.</summary>
        private void SubstituteFields()
        {
            var total = _totalPages > 0 ? _totalPages : _result.Pages.Count;
            for (var i = 0; i < _result.Pages.Count; i++)
            {
                var page = _result.Pages[i];
                var props = i < _pageSections.Count ? _pageSections[i] : _props;
                foreach (var op in page.Ops)
                {
                    var text = op as TextOp;
                    if (text == null || text.Field == FieldKind.None)
                        continue;
                    text.Text = text.Field == FieldKind.NumPages
                        ? total.ToString(CultureInfo.InvariantCulture)
                        : FormatPageNumber(page.Number, props);
                    text.Field = FieldKind.None;
                }
            }
        }
    }
}
