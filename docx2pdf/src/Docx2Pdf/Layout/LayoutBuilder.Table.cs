using System;
using System.Collections.Generic;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    internal sealed partial class LayoutBuilder
    {
        /// <summary>Maximum height a single fragment may occupy before it is split across pages.</summary>
        public double MaxBlockHeight = 10000;

        private sealed class CellPlan
        {
            public TableCell Cell;
            public double X;
            public double Width;
            public double MarginLeft, MarginRight, MarginTop, MarginBottom;
            public List<Fragment> Fragments = new List<Fragment>();
            public int Position;
            public double ContentHeight;
            public int ColumnIndex;
            public int ColumnSpan;
            /// <summary>The vertical merge this cell belongs to continues in the next row.</summary>
            public bool MergeContinuesBelow;
            /// <summary>Height of the rows the merge spans below this one (vertical centring).</summary>
            public double SpanBelowPt;
        }

        public List<Fragment> BuildTable(Table table, double availableWidth)
        {
            var fragments = new List<Fragment>();
            if (table.Rows.Count == 0)
                return fragments;

            var columns = ResolveColumns(table, availableWidth);
            var tableWidth = 0.0;
            foreach (var w in columns)
                tableWidth += w;

            var offset = table.IndentPt;
            // Percent-width tables align their first-cell text with the indent; the
            // border sits one cell margin further left (see the target computation).
            if (table.PreferredWidthIsPercent && table.Alignment == TextAlignment.Left && !_inTableCell)
                offset -= table.CellMarginLeftPt;
            if (table.Alignment == TextAlignment.Center)
            {
                // A centred table wider than the text column hangs into both margins evenly,
                // so the offset may go negative (Word does the same).
                offset = (availableWidth - tableWidth) / 2;
            }
            else if (table.Alignment == TextAlignment.Right)
            {
                offset = Math.Max(0, availableWidth - tableWidth);
            }
            // Word honours the stated indent even when the table then hangs into the right
            // margin (SaBC Features table: tblInd 607tw + grid 8886tw > 9360tw body — Word
            // keeps the table at the indent).

            var tableId = new object();

            var rowPlans = new List<List<CellPlan>>();
            foreach (var row in table.Rows)
                rowPlans.Add(PlanRow(table, row, columns, offset));

            var extraHeight = DistributeMergedCellHeights(table, rowPlans);

            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var row = table.Rows[rowIndex];
                var plans = rowPlans[rowIndex];
                if (plans.Count == 0)
                    continue;

                var isLastRow = rowIndex == table.Rows.Count - 1;
                var nextPlans = isLastRow ? null : rowPlans[rowIndex + 1];
                var source = new TableRowSource(table, row, plans, nextPlans, rowIndex, isLastRow, tableId,
                                               extraHeight[rowIndex], _ctx.LegacyCellSpacing);
                var fragment = source.Next(source.NaturalHeight, true);
                if (fragment == null)
                    continue;
                fragment.RowSource = source;
                fragments.Add(fragment);
            }

            return fragments;
        }

        /// <summary>
        /// A vertically merged cell's content has to fit across all the rows it spans.
        /// Returns, per row, the extra height needed so merged content is not clipped.
        /// </summary>
        private static double[] DistributeMergedCellHeights(Table table, List<List<CellPlan>> rowPlans)
        {
            var extra = new double[table.Rows.Count];
            var natural = new double[table.Rows.Count];

            for (var i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                double content = 0, margins = 0;
                foreach (var plan in rowPlans[i])
                {
                    if (plan.Cell.VMerge == VerticalMerge.Restart)
                        continue;                       // measured across the whole span below
                    content = Math.Max(content, plan.ContentHeight);
                    margins = Math.Max(margins, plan.MarginTop + plan.MarginBottom);
                }
                natural[i] = content + margins;
                if (row.HeightPt > 0)
                    natural[i] = row.HeightExact ? row.HeightPt : Math.Max(natural[i], row.HeightPt);
            }

            for (var i = 0; i < table.Rows.Count; i++)
            {
                foreach (var plan in rowPlans[i])
                {
                    if (plan.Cell.VMerge != VerticalMerge.Restart)
                        continue;

                    // How far down does this merge continue?
                    var last = i;
                    for (var j = i + 1; j < table.Rows.Count; j++)
                    {
                        CellPlan continuation = null;
                        foreach (var other in rowPlans[j])
                        {
                            if (other.ColumnIndex == plan.ColumnIndex && other.Cell.VMerge == VerticalMerge.Continue)
                            {
                                continuation = other;
                                break;
                            }
                        }
                        if (continuation == null)
                            break;
                        // Every cell of the merge except the bottom one keeps its bottom
                        // edge open (the border through a merged cell is not drawn).
                        foreach (var other in rowPlans[j - 1])
                        {
                            if (other.ColumnIndex == plan.ColumnIndex)
                                other.MergeContinuesBelow = true;
                        }
                        last = j;
                    }

                    var needed = plan.ContentHeight + plan.MarginTop + plan.MarginBottom;
                    double available = 0;
                    for (var j = i; j <= last; j++)
                        available += natural[j] + extra[j];

                    if (needed > available)
                        extra[last] += needed - available;

                    // Height the merge spans below its first row, for vertical centring
                    // of the restart cell's content across the whole span.
                    double below = 0;
                    for (var j = i + 1; j <= last; j++)
                        below += natural[j] + extra[j];
                    plan.SpanBelowPt = below;
                }
            }

            return extra;
        }

        private List<double> ResolveColumns(Table table, double availableWidth)
        {
            var columns = new List<double>(table.Grid);

            var maxCells = 0;
            foreach (var row in table.Rows)
            {
                var span = 0;
                foreach (var cell in row.Cells)
                    span += Math.Max(1, cell.GridSpan);
                if (span > maxCells)
                    maxCells = span;
            }

            if (columns.Count < maxCells)
            {
                // Derive missing columns from declared cell widths, or split the remainder
                // evenly. Rows whose cells carry no widths at all are skipped — a generated
                // document may declare widths on its header rows only (the MPOB licence's
                // 1008-row lot table: 1005 width-less rows, widths on the 3 header rows).
                var widths = new List<double>();
                foreach (var row in table.Rows)
                {
                    if (row.Cells.Count == 0)
                        continue;
                    var span = 0;
                    foreach (var cell in row.Cells)
                        span += Math.Max(1, cell.GridSpan);
                    if (span != maxCells)
                        continue;
                    var any = false;
                    widths.Clear();
                    foreach (var cell in row.Cells)
                    {
                        var each = Math.Max(1, cell.GridSpan);
                        var w = cell.WidthIsPercent ? availableWidth * cell.WidthPt / 100.0 : cell.WidthPt;
                        if (w > 0)
                            any = true;
                        for (var i = 0; i < each; i++)
                            widths.Add(w / each);
                    }
                    if (any)
                        break;
                    widths.Clear();
                }
                while (columns.Count < maxCells)
                    columns.Add(columns.Count < widths.Count ? widths[columns.Count] : 0);
            }

            // A grid-less table auto-fits: Word grows a column past its declared width when
            // unwrapped cell content needs it (the MPOB licence's lot table declares a
            // 120pt Mukim column but renders ~124.5pt so "BATU API LAND DISTRICT" stays on
            // one line). Oversized results are pulled back by the width cap below.
            if (table.Grid.Count == 0 && maxCells > 0)
            {
                var prefs = new double[maxCells];
                foreach (var row in table.Rows)
                {
                    var column = 0;
                    foreach (var cell in row.Cells)
                    {
                        var span = Math.Max(1, cell.GridSpan);
                        if (span == 1 && column < maxCells)
                        {
                            var margins = (cell.MarginLeftPt ?? table.CellMarginLeftPt)
                                          + (cell.MarginRightPt ?? table.CellMarginRightPt);
                            foreach (var block in cell.Blocks)
                            {
                                var paragraph = block as Paragraph;
                                if (paragraph == null)
                                    continue;
                                var w = MaxLineWidth(paragraph) + margins;
                                if (w > prefs[column])
                                    prefs[column] = w;
                            }
                        }
                        column += span;
                    }
                }
                for (var i = 0; i < columns.Count && i < maxCells; i++)
                    columns[i] = Math.Max(columns[i], prefs[i]);
            }

            var total = 0.0;
            foreach (var w in columns)
                total += w;

            var target = total;
            if (table.PreferredWidthPt.HasValue && table.PreferredWidthPt.Value > 0)
            {
                // A percent table with a complete tblGrid renders that grid VERBATIM:
                // the grid is the column layout Word itself computed for the percent
                // width when the document was saved (sample1: City 100% = 478.8pt
                // grid = margin−5.4..margin+5.4; p4 outer 70% = its 335.15pt grid,
                // NOT 70%+margins; p4 inner 80% = its 130.5pt grid). Without a full
                // grid, fall back to percent-of-available; the border overhang by the
                // outer cell margins applies to top-level tables only.
                if (table.PreferredWidthIsPercent)
                {
                    if (!(table.Grid.Count >= maxCells && total > 0))
                        target = availableWidth * table.PreferredWidthPt.Value / 100.0
                                 + (_inTableCell ? 0 : table.CellMarginLeftPt + table.CellMarginRightPt);
                }
                else
                {
                    target = table.PreferredWidthPt.Value;
                }
            }
            if (target <= 0)
                target = availableWidth;
            // Word honours a declared width moderately wider than the text column (the table
            // hangs into the margins); an oversized PREFERRED width is pulled back to fit —
            // but a GRID that is itself wider than the body renders verbatim: Word lets the
            // table overflow the right margin and paper edge instead of shrinking columns
            // (MyLesen FEE03 title block: grid 25294tw on a ~14850tw landscape body — the
            // value column keeps its 538pt and the last column runs off the page).
            if (target > availableWidth + 36 && total <= availableWidth + 36)
                target = availableWidth;

            if (total <= 0)
            {
                var each = target / Math.Max(1, columns.Count);
                for (var i = 0; i < columns.Count; i++)
                    columns[i] = each;
            }
            else if (Math.Abs(total - target) > 0.5)
            {
                var scale = target / total;
                for (var i = 0; i < columns.Count; i++)
                    columns[i] *= scale;
            }

            if (columns.Count == 0)
                columns.Add(availableWidth);
            return columns;
        }

        /// <summary>Width of the paragraph's longest line when nothing wraps.</summary>
        private double MaxLineWidth(Paragraph paragraph)
        {
            var atoms = BuildAtoms(paragraph);
            double max = 0, current = 0;
            foreach (var atom in atoms)
            {
                if (atom is BreakAtom)
                {
                    if (current > max)
                        max = current;
                    current = 0;
                    continue;
                }
                current += atom.Width;
            }
            return Math.Max(max, current);
        }

        private List<CellPlan> PlanRow(Table table, TableRow row, List<double> columns, double offset)
        {
            var plans = new List<CellPlan>();
            var column = 0;
            var x = offset;

            foreach (var cell in row.Cells)
            {
                var span = Math.Max(1, cell.GridSpan);
                double width = 0;
                for (var i = 0; i < span && column + i < columns.Count; i++)
                    width += columns[column + i];
                if (width <= 0)
                    width = columns.Count > 0 ? columns[Math.Min(column, columns.Count - 1)] : 72;

                var plan = new CellPlan
                {
                    Cell = cell,
                    X = x,
                    Width = width,
                    ColumnIndex = column,
                    ColumnSpan = span,
                    MarginLeft = cell.MarginLeftPt ?? table.CellMarginLeftPt,
                    MarginRight = cell.MarginRightPt ?? table.CellMarginRightPt,
                    MarginTop = cell.MarginTopPt ?? table.CellMarginTopPt,
                    MarginBottom = cell.MarginBottomPt ?? table.CellMarginBottomPt,
                };

                var inner = plan.Width - plan.MarginLeft - plan.MarginRight;
                if (inner < 8)
                {
                    plan.MarginLeft = plan.MarginRight = 1;
                    inner = Math.Max(4, plan.Width - 2);
                }

                if (cell.VMerge != VerticalMerge.Continue)
                {
                    // Word counts the last paragraph's space-after inside the cell, so nothing is trimmed here.
                    var wasInCell = _inTableCell;
                    _inTableCell = true;
                    try
                    {
                        plan.Fragments = BuildBlocks(cell.Blocks, inner);
                    }
                    finally
                    {
                        _inTableCell = wasInCell;
                    }
                    foreach (var fragment in plan.Fragments)
                        plan.ContentHeight += fragment.Height;
                }

                plans.Add(plan);
                column += span;
                x += width;
            }

            return plans;
        }

        /// <summary>Emits one table row, in slices when it does not fit on the remaining page.</summary>
        private sealed class TableRowSource : ITableRowSource
        {
            private readonly Table _table;
            private readonly TableRow _row;
            private readonly List<CellPlan> _plans;
            private readonly int _rowIndex;
            private readonly bool _isLastRow;
            private readonly object _tableId;
            private readonly double _marginHeight;
            private bool _first = true;

            public TableRowSource(Table table, TableRow row, List<CellPlan> plans, List<CellPlan> nextPlans,
                                  int rowIndex, bool isLastRow, object tableId, double extraHeight,
                                  bool legacyEmptyRowCollapse)
            {
                _table = table;
                _row = row;
                _plans = plans;
                _rowIndex = rowIndex;
                _isLastRow = isLastRow;
                _tableId = tableId;
                _extraHeight = extraHeight;

                foreach (var plan in plans)
                    _marginHeight = Math.Max(_marginHeight, plan.MarginTop + plan.MarginBottom);

                // Horizontal table borders occupy vertical layout space in Word (PROBE13 via
                // COM: the header table takes content + both border widths; the repeated
                // Features header row renders 26.25pt for 25.6pt of content; p11's cut
                // rejects a line that would only fit without the border). The row box is
                // top border (first row only — an insideH boundary belongs to the row above)
                // + content + bottom border.
                foreach (var plan in plans)
                {
                    var cellBorders = plan.Cell.Borders;
                    var tableBorders = table.Borders;
                    if (rowIndex == 0)
                    {
                        var top = Pick(cellBorders == null ? null : cellBorders.Top,
                                       tableBorders == null ? null : tableBorders.Top);
                        if (top != null && top.IsVisible)
                            _topBorderPt = Math.Max(_topBorderPt, top.WidthPt);
                    }
                    var bottom = Pick(cellBorders == null ? null : cellBorders.Bottom,
                                      tableBorders == null ? null : (isLastRow ? tableBorders.Bottom : tableBorders.InsideH));
                    if (bottom != null && bottom.IsVisible)
                        _bottomBorderPt = Math.Max(_bottomBorderPt, bottom.WidthPt);
                }
                // The boundary below this row is resolved against the NEXT row's TOP
                // borders too — a conditional-region border lives on the cells below it
                // (sample1 College table: the lastRow double sits on the Total row's
                // cells' Top; Word's Elm row is 24.75pt = 22.5 + that 2.25pt band).
                if (nextPlans != null)
                {
                    foreach (var plan in nextPlans)
                    {
                        var top = Pick(plan.Cell.Borders == null ? null : plan.Cell.Borders.Top,
                                       table.Borders == null ? null : table.Borders.InsideH);
                        if (top != null && top.IsVisible)
                            _bottomBorderPt = Math.Max(_bottomBorderPt, top.WidthPt);
                    }
                }

                // A cell that starts a vertical merge is measured across the rows it spans,
                // so it does not force this row to be tall on its own.
                var content = 0.0;
                foreach (var plan in plans)
                {
                    if (plan.Cell.VMerge == VerticalMerge.Restart)
                        continue;
                    content = Math.Max(content, plan.ContentHeight);
                }
                var height = content + _marginHeight;
                if (row.HeightPt > 0)
                    height = row.HeightExact ? row.HeightPt : Math.Max(height, row.HeightPt);
                // No minimum floor: Word renders rows at their content height even when
                // tiny (licence form COM: the sz-8-mark spacer rows measure 4.5pt and
                // the 8pt-mark row 8.25 — an earlier 8pt floor inflated the former, and
                // an "empty rows collapse to trHeight" probe broke the latter; both are
                // simply the MARK LINE heights).
                NaturalHeight = Math.Max(1, height + extraHeight + _topBorderPt + _bottomBorderPt);

                if (Environment.GetEnvironmentVariable("DOCX2PDF_DEBUG_ROWS") != null)
                {
                    var parts = new System.Text.StringBuilder();
                    foreach (var plan in plans)
                    {
                        parts.Append(plan.ContentHeight.ToString("F1"));
                        parts.Append('[');
                        foreach (var f in plan.Fragments)
                            parts.Append(f.Height.ToString("F1")).Append(' ');
                        parts.Append("] ");
                    }
                    string label = null;
                    foreach (var plan in plans)
                    {
                        foreach (var block in plan.Cell.Blocks)
                        {
                            var para = block as Paragraph;
                            if (para == null)
                                continue;
                            var text = PlainText(para);
                            if (!string.IsNullOrEmpty(text))
                            {
                                label = text.Length > 28 ? text.Substring(0, 28) : text;
                                break;
                            }
                        }
                        if (label != null)
                            break;
                    }
                    Console.Error.WriteLine("row: trH={0:F1} content={1:F1} margins={2:F1} => {3:F1}  cells: {4} '{5}'",
                        row.HeightPt, content, _marginHeight, NaturalHeight, parts, label ?? string.Empty);
                }
            }

            private readonly double _extraHeight;
            private double _topBorderPt;
            private double _bottomBorderPt;

            public double NaturalHeight { get; private set; }

            public bool CantSplit
            {
                get { return _row.CantSplit || _row.IsHeader || _row.HeightExact; }
            }

            public bool Exhausted
            {
                get
                {
                    foreach (var plan in _plans)
                    {
                        if (plan.Position < plan.Fragments.Count)
                            return false;
                    }
                    return true;
                }
            }

            public void Rewind()
            {
                foreach (var plan in _plans)
                    plan.Position = 0;
                _first = true;
            }

            public Fragment Next(double maxHeight, bool forceProgress)
            {
                if (Exhausted && !_first)
                    return null;

                // A trHeight "atLeast" minimum is a per-page minimum: a split row never
                // starts in less space than it (MyLesen p10: row 5 had 79pt free under
                // row 4 but trHeight 135.4 — Word moves the whole row to p11 even though
                // every cell could have placed lines). At the top of a page the row
                // starts regardless, else a minimum taller than the page would never place.
                if (_first && !forceProgress && _row.HeightPt > 0 && !_row.HeightExact
                    && maxHeight < _row.HeightPt + _topBorderPt + _bottomBorderPt - 0.001)
                    return null;

                // The first slice carries the table's top border; every slice carries the
                // border that closes it (the row's bottom border or the cut). Both take
                // space away from the cell content.
                var topBorder = _first ? _topBorderPt : 0;
                var available = Math.Max(8, maxHeight - _marginHeight - topBorder - _bottomBorderPt);

                // Snapshot for the start-of-row check below: positions are rolled back when
                // the row turns out not to start on this page after all.
                int[] startPositions = null;
                if (_first && !forceProgress)
                {
                    startPositions = new int[_plans.Count];
                    for (var i = 0; i < _plans.Count; i++)
                        startPositions[i] = _plans[i].Position;
                }

                var taken = new List<List<Fragment>>();
                var partHeight = 0.0;
                var anyTaken = false;

                foreach (var plan in _plans)
                {
                    // A cell that starts a vertical merge renders across the rows it spans, so its
                    // content is emitted in full here; the span's total height accommodates it.
                    var unbounded = plan.Cell.VMerge == VerticalMerge.Restart;

                    var list = new List<Fragment>();
                    var height = 0.0;
                    while (plan.Position < plan.Fragments.Count)
                    {
                        var fragment = plan.Fragments[plan.Position];
                        // With forceProgress the first fragment is always taken, so a slice is
                        // never empty. The epsilon absorbs float noise between the natural
                        // height computed at construction and the available height derived
                        // back from it — without it a row whose content exactly fills it can
                        // silently drop its last line (Xu portfolio p10, "Shanghai University").
                        if (!unbounded && height + fragment.Height > available + 0.001 && !(list.Count == 0 && forceProgress))
                            break;
                        list.Add(fragment);
                        height += fragment.Height;
                        plan.Position++;
                        anyTaken = true;
                    }

                    // Word's cut rules at a row split (PROBE8 variant B, confirmed by PROBE10):
                    //  - a paragraph may only COMPLETE in a slice if its trailing boundary
                    //    spacing also fits; otherwise its last line moves to the next slice
                    //    even though the line itself would fit.
                    if (!unbounded && plan.Position < plan.Fragments.Count && list.Count > 1)
                    {
                        var next = plan.Fragments[plan.Position];
                        var last = list[list.Count - 1];
                        if (next.IsSpacing && !next.IsSpaceBefore && !last.IsSpacing
                            && last.ParagraphId != null && last.LineIndex == last.LineCount - 1
                            && ReferenceEquals(next.ParagraphId, last.ParagraphId))
                        {
                            height -= last.Height;
                            list.RemoveAt(list.Count - 1);
                            plan.Position--;
                        }
                    }

                    // Widow/orphan control at the cut (PROBE10). PROBE8's earlier "no
                    // pull-back" observation came from legacy-layout Word: that probe had no
                    // settings.xml. A compatibilityMode-15 document pulls lines back at a row
                    // split exactly as at a page break, and w:widowControl val="0" turns it
                    // off (PROBE10 N-variants):
                    //  - widow: the remainder of a split paragraph is never a lone last
                    //    line — a second line moves with it (SaBC p6 splits 2/2, not 3/1);
                    //  - orphan: a lone first line never stays at the bottom of a slice —
                    //    it moves to the next slice together with its space-before.
                    if (!unbounded && plan.Position < plan.Fragments.Count && list.Count > 0)
                    {
                        var next = plan.Fragments[plan.Position];
                        var last = list[list.Count - 1];
                        if (last.WidowControl && !last.IsSpacing && last.ParagraphId != null
                            && ReferenceEquals(next.ParagraphId, last.ParagraphId)
                            && last.LineIndex < last.LineCount - 1)
                        {
                            var pull = 0;
                            if (last.LineIndex == 0)
                            {
                                pull = 1;
                                while (pull < list.Count
                                       && list[list.Count - 1 - pull].IsSpacing
                                       && ReferenceEquals(list[list.Count - 1 - pull].ParagraphId, last.ParagraphId))
                                    pull++;
                            }
                            else if (last.LineIndex == last.LineCount - 2)
                            {
                                if (last.LineIndex >= 2)
                                {
                                    pull = 1;
                                }
                                else
                                {
                                    // Too short to split two/two: pulling one line back would
                                    // orphan the first, so the WHOLE paragraph moves — Word's
                                    // classic rule for three-line paragraphs (MyLesen p9:
                                    // item d lands complete at the top of p10).
                                    pull = last.LineIndex + 1;
                                    while (pull < list.Count
                                           && list[list.Count - 1 - pull].IsSpacing
                                           && ReferenceEquals(list[list.Count - 1 - pull].ParagraphId, last.ParagraphId))
                                        pull++;
                                }
                            }
                            for (var k = 0; k < pull && list.Count > 1; k++)
                            {
                                var pulled = list[list.Count - 1];
                                height -= pulled.Height;
                                list.RemoveAt(list.Count - 1);
                                plan.Position--;
                            }
                        }
                    }
                    taken.Add(list);
                    // Content of a merged cell spans several rows, so it does not size this one.
                    if (height > partHeight && plan.Cell.VMerge != VerticalMerge.Restart)
                        partHeight = height;
                }

                // A row only starts on a page where every cell can place real content: when a
                // cell's first line or image does not fit and only spacing was taken, Word
                // moves the whole row to the next page rather than emit a label-only sliver
                // (p7: the Quick Statistics row waits for p8 because its screenshot does not
                // fit under the Knowledge Explorer row).
                if (startPositions != null)
                {
                    for (var i = 0; i < _plans.Count; i++)
                    {
                        var plan = _plans[i];
                        var tookContent = false;
                        foreach (var f in taken[i])
                        {
                            if (!f.IsSpacing)
                            {
                                tookContent = true;
                                break;
                            }
                        }
                        if (tookContent)
                            continue;
                        var pendingContent = false;
                        for (var j = plan.Position; j < plan.Fragments.Count; j++)
                        {
                            if (!plan.Fragments[j].IsSpacing)
                            {
                                pendingContent = true;
                                break;
                            }
                        }
                        if (pendingContent)
                        {
                            for (var k = 0; k < _plans.Count; k++)
                                _plans[k].Position = startPositions[k];
                            return null;
                        }
                    }
                }

                if (!anyTaken && !_first)
                    return null;

                // The first slice honours the row's natural height (trHeight minimums, merged
                // cells) but never past the space the page offered. A slice that continues on
                // the next page hugs its content instead — Word closes the cut right under the
                // last kept line (p6: cut at 901px, page bottom 947px) — unless a vertically
                // merged cell still needs the full offered space.
                var exhausted = Exhausted;
                var needsFill = exhausted;
                foreach (var plan in _plans)
                {
                    if (plan.Cell.VMerge == VerticalMerge.Restart)
                        needsFill = true;
                }
                var rowHeight = partHeight + _marginHeight + topBorder + _bottomBorderPt;
                if (_first && needsFill)
                    rowHeight = Math.Max(rowHeight, Math.Min(NaturalHeight, maxHeight));
                if (_row.HeightPt > 0)
                {
                    var minHeight = _row.HeightPt + topBorder + _bottomBorderPt;
                    if (_row.HeightExact)
                    {
                        if (_first && exhausted)
                            rowHeight = minHeight;
                    }
                    else
                    {
                        // trHeight "atLeast" pads EVERY slice of a split row, not just the
                        // whole row: Word's MyLesen p10 remainder measures exactly trHeight
                        // 237.8 + bottom border for 225.5pt of content — the slice keeps the
                        // row minimum on each page it appears, capped at the space offered.
                        rowHeight = Math.Max(rowHeight, Math.Min(minHeight, maxHeight));
                    }
                }
                if (rowHeight <= 0)
                    rowHeight = 8;

                var fragment2 = new Fragment(rowHeight) { TableId = _tableId };
                fragment2.EdgeExtentPt = topBorder + _bottomBorderPt;
                if (_row.IsHeader)
                    fragment2.IsTableHeaderRow = true;

                foreach (var plan in _plans)
                {
                    var shading = plan.Cell.Shading ?? _table.Shading;
                    if (shading.HasValue)
                    {
                        fragment2.Add(new RectOp
                        {
                            X = plan.X, Y = 0, Width = plan.Width, Height = rowHeight, Color = shading.Value,
                        });
                    }
                }

                for (var i = 0; i < _plans.Count; i++)
                {
                    var plan = _plans[i];
                    var content = taken[i];
                    double used = 0;
                    foreach (var f in content)
                        used += f.Height;

                    var top = plan.MarginTop + topBorder;
                    // A cell that starts a vertical merge aligns its content across the
                    // whole span, not just this row (Nota table: SLGTISU centres over the
                    // two rows its merge covers).
                    var free = rowHeight - _marginHeight - topBorder - _bottomBorderPt - used
                               + plan.SpanBelowPt;
                    if (free > 0)
                    {
                        if (plan.Cell.VAlign == VerticalCellAlignment.Center)
                            top += free / 2;
                        else if (plan.Cell.VAlign == VerticalCellAlignment.Bottom)
                            top += free;
                    }

                    var y = top;
                    foreach (var f in content)
                    {
                        var firstOp = fragment2.Ops.Count;
                        fragment2.AddTranslated(f.Ops, plan.X + plan.MarginLeft, y);
                        ClipToCell(fragment2.Ops, firstOp, plan.X + plan.Width);
                        if (f.HeadingLevel > 0 && fragment2.HeadingLevel == 0)
                        {
                            fragment2.HeadingLevel = f.HeadingLevel;
                            fragment2.HeadingText = f.HeadingText;
                        }
                        y += f.Height;
                    }
                }

                EmitRowBorders(fragment2, _table, _row, _plans, rowHeight, _rowIndex,
                               _isLastRow && exhausted, _first, exhausted,
                               topBorder, _bottomBorderPt);
                _first = false;
                return fragment2;
            }
        }

        /// <summary>
        /// Content wider than its cell is clipped at the cell's right border: Word cuts an
        /// oversized inline picture there instead of painting over the neighbouring column
        /// (MyLesen p39: the licence-table screenshot ends exactly at the Description cell
        /// edge). Pictures clip via their source crop; frame lines and rects are clamped.
        /// Floating (anchored) content legitimately escapes the cell and is left alone.
        /// </summary>
        private static void ClipToCell(List<DrawOp> ops, int from, double right)
        {
            for (var i = from; i < ops.Count; i++)
            {
                var image = ops[i] as ImageOp;
                if (image != null && image.RotationDeg == 0 && image.X + image.Width > right + 0.1)
                {
                    var keep = (right - image.X) / image.Width;
                    if (keep <= 0)
                    {
                        ops.RemoveAt(i);
                        i--;
                        continue;
                    }
                    var span = 1 - image.CropLeft - image.CropRight;
                    image.CropRight = 1 - image.CropLeft - span * keep;
                    image.Width = right - image.X;
                    continue;
                }
                var line = ops[i] as LineOp;
                if (line != null && (line.X1 > right + 0.1 || line.X2 > right + 0.1))
                {
                    if (line.X1 > right && line.X2 > right)
                    {
                        ops.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (line.X1 > right) line.X1 = right;
                    if (line.X2 > right) line.X2 = right;
                    continue;
                }
                var rect = ops[i] as RectOp;
                if (rect != null && rect.X + rect.Width > right + 0.1)
                    rect.Width = Math.Max(0, right - rect.X);
            }
        }

        private static void EmitRowBorders(Fragment fragment, Table table, TableRow row, List<CellPlan> plans,
                                           double rowHeight, int rowIndex, bool isLastRow, bool isFirstPart,
                                           bool isLastPart, double topBorderPt, double bottomBorderPt)
        {
            var tableBorders = table.Borders;

            for (var i = 0; i < plans.Count; i++)
            {
                var plan = plans[i];
                var cellBorders = plan.Cell.Borders;
                var isFirstColumn = plan.ColumnIndex == 0;
                var isLastColumn = i == plans.Count - 1;
                var isFirstRow = rowIndex == 0;

                var top = Pick(cellBorders == null ? null : cellBorders.Top,
                               tableBorders == null ? null : (isFirstRow ? tableBorders.Top : tableBorders.InsideH));
                var bottom = Pick(cellBorders == null ? null : cellBorders.Bottom,
                                  tableBorders == null ? null : (isLastRow ? tableBorders.Bottom : tableBorders.InsideH));
                var left = Pick(cellBorders == null ? null : cellBorders.Left,
                                tableBorders == null ? null : (isFirstColumn ? tableBorders.Left : tableBorders.InsideV));
                var right = Pick(cellBorders == null ? null : cellBorders.Right,
                                 tableBorders == null ? null : (isLastColumn ? tableBorders.Right : tableBorders.InsideV));

                // A continued vertical merge keeps the cell open at the top, and any cell
                // whose merge continues below keeps its bottom edge open — no border is
                // drawn through the middle of a merged cell.
                if (plan.Cell.VMerge == VerticalMerge.Continue)
                    top = null;
                if (plan.MergeContinuesBelow)
                    bottom = null;

                // Border lines sit inside the space the row box reserves for them.
                if (top != null && top.IsVisible && (isFirstPart || isFirstRow))
                    fragment.Add(HorizontalLine(plan.X, plan.X + plan.Width, topBorderPt / 2, top));
                // The bottom edge is also drawn when the row continues on the next page:
                // Word closes a cut slice with the row's bottom border (PROBE10, SaBC p6).
                if (bottom != null && bottom.IsVisible)
                    fragment.Add(HorizontalLine(plan.X, plan.X + plan.Width, rowHeight - bottomBorderPt / 2, bottom));
                if (left != null && left.IsVisible)
                    fragment.Add(new LineOp
                    {
                        X1 = plan.X, Y1 = 0, X2 = plan.X, Y2 = rowHeight,
                        Width = left.WidthPt, Color = left.Color, Style = left.Style,
                    });
                if (right != null && right.IsVisible)
                    fragment.Add(new LineOp
                    {
                        X1 = plan.X + plan.Width, Y1 = 0, X2 = plan.X + plan.Width, Y2 = rowHeight,
                        Width = right.WidthPt, Color = right.Color, Style = right.Style,
                    });
            }
        }

        private static Border Pick(Border cell, Border table)
        {
            if (cell != null && cell.Style != BorderStyle.None)
                return cell;
            if (cell != null && cell.Style == BorderStyle.None)
                return null;      // explicitly switched off on the cell
            return table;
        }
    }
}
