using System;
using System.Collections.Generic;
using Docx2Pdf.Fonts;
using Docx2Pdf.Images;
using Docx2Pdf.Model;

namespace Docx2Pdf.Layout
{
    /// <summary>
    /// A drawing primitive positioned in a top-left origin coordinate system measured in points.
    /// The renderer flips Y when writing the PDF content stream.
    /// </summary>
    internal abstract class DrawOp
    {
        /// <summary>Painted after all body content: floating frames sit above inline images.</summary>
        public bool Overlay;

        public abstract DrawOp Translate(double dx, double dy);
    }

    internal sealed class TextOp : DrawOp
    {
        public double X;
        public double Y;                 // baseline
        public PdfFontBase Font;
        public double Size;
        public uint Color;
        public string Text;
        public double CharSpacing;
        /// <summary>When set, the renderer substitutes the live page number.</summary>
        public FieldKind Field;

        public override DrawOp Translate(double dx, double dy)
        {
            return new TextOp
            {
                X = X + dx, Y = Y + dy, Font = Font, Size = Size, Color = Color,
                Text = Text, CharSpacing = CharSpacing, Field = Field,
            };
        }
    }

    internal sealed class RectOp : DrawOp
    {
        public double X, Y, Width, Height;
        public uint Color;

        public override DrawOp Translate(double dx, double dy)
        {
            return new RectOp { X = X + dx, Y = Y + dy, Width = Width, Height = Height, Color = Color };
        }
    }

    internal sealed class LineOp : DrawOp
    {
        public double X1, Y1, X2, Y2;
        public uint Color;
        public double Width = 0.5;
        public BorderStyle Style = BorderStyle.Single;

        public override DrawOp Translate(double dx, double dy)
        {
            return new LineOp
            {
                X1 = X1 + dx, Y1 = Y1 + dy, X2 = X2 + dx, Y2 = Y2 + dy,
                Color = Color, Width = Width, Style = Style,
            };
        }
    }

    internal sealed class ImageOp : DrawOp
    {
        public double X, Y, Width, Height;
        public string Key;
        public DecodedImage Image;
        public double RotationDeg;
        /// <summary>Source crop fractions 0..1 per edge (a:srcRect).</summary>
        public double CropLeft, CropTop, CropRight, CropBottom;

        public override DrawOp Translate(double dx, double dy)
        {
            return new ImageOp
            {
                X = X + dx, Y = Y + dy, Width = Width, Height = Height,
                Key = Key, Image = Image, RotationDeg = RotationDeg,
                CropLeft = CropLeft, CropTop = CropTop, CropRight = CropRight, CropBottom = CropBottom,
            };
        }
    }

    internal sealed class LinkOp : DrawOp
    {
        public double X, Y, Width, Height;
        public string Url;
        public string Anchor;

        public override DrawOp Translate(double dx, double dy)
        {
            return new LinkOp { X = X + dx, Y = Y + dy, Width = Width, Height = Height, Url = Url, Anchor = Anchor };
        }
    }

    /// <summary>
    /// A floating frame (anchored picture or text box). X/Y mark the anchor point in the text flow;
    /// the page composer resolves the final position from the anchor's reference frames.
    /// </summary>
    internal sealed class AnchoredOp : DrawOp
    {
        public double X, Y;
        public double Width, Height;
        /// <summary>Frame content, positioned relative to the frame's top-left corner.</summary>
        public List<DrawOp> Content = new List<DrawOp>();

        public AnchorRelativeFrom HorizontalFrom;
        public double HorizontalOffset;
        public string HorizontalAlign;
        public AnchorRelativeFrom VerticalFrom;
        public double VerticalOffset;
        public string VerticalAlign;
        public bool BehindDoc;

        public override DrawOp Translate(double dx, double dy)
        {
            return new AnchoredOp
            {
                X = X + dx, Y = Y + dy, Width = Width, Height = Height, Content = Content,
                HorizontalFrom = HorizontalFrom, HorizontalOffset = HorizontalOffset, HorizontalAlign = HorizontalAlign,
                VerticalFrom = VerticalFrom, VerticalOffset = VerticalOffset, VerticalAlign = VerticalAlign,
                BehindDoc = BehindDoc,
            };
        }
    }

    internal sealed class BookmarkOp : DrawOp
    {
        public double X, Y;
        public string Name;

        public override DrawOp Translate(double dx, double dy)
        {
            return new BookmarkOp { X = X + dx, Y = Y + dy, Name = Name };
        }
    }

    /// <summary>
    /// A table row that can be rendered in several vertical slices, so a tall row
    /// continues onto the next page the way Word splits rows.
    /// </summary>
    internal interface ITableRowSource
    {
        /// <summary>Full height of the row when nothing needs to be split.</summary>
        double NaturalHeight { get; }
        /// <summary>True when the row must stay on one page (w:cantSplit).</summary>
        bool CantSplit { get; }
        /// <summary>True once every cell has been emitted.</summary>
        bool Exhausted { get; }
        /// <summary>Restarts emission from the top of the row.</summary>
        void Rewind();
        /// <summary>
        /// Emits the next slice of at most <paramref name="maxHeight"/>, or null when nothing fits.
        /// With <paramref name="forceProgress"/> each cell contributes at least one fragment.
        /// </summary>
        Fragment Next(double maxHeight, bool forceProgress);
    }

    /// <summary>An atomic vertical slice of content (one line, one table row part, one spacing gap).</summary>
    internal sealed class Fragment
    {
        public double Height;
        public List<DrawOp> Ops = new List<DrawOp>();
        /// <summary>Keeps this fragment on the same page as the following one when possible.</summary>
        public bool KeepWithNext;
        public bool PageBreakBefore;
        /// <summary>
        /// Set when the paragraph ends with an explicit page break: the line itself stays on the
        /// current page (Word keeps the paragraph mark there) and following content starts a new page.
        /// </summary>
        public bool PageBreakAfter;
        /// <summary>Pure spacing, dropped when it lands at the top of a page.</summary>
        public bool IsSpacing;
        /// <summary>
        /// Space-before of the paragraph that follows. Word keeps a paragraph's own space-before
        /// at the top of a page; only spilled-over space-after is dropped there.
        /// </summary>
        public bool IsSpaceBefore;
        /// <summary>Heading information used to build the PDF outline.</summary>
        public int HeadingLevel;
        public string HeadingText;
        /// <summary>Repeated at the top of every page a table continues onto.</summary>
        public bool IsTableHeaderRow;
        /// <summary>Identity of the table this fragment belongs to, used to repeat header rows.</summary>
        public object TableId;
        /// <summary>Set on table row fragments that can be split across pages.</summary>
        public ITableRowSource RowSource;
        /// <summary>Identity of the paragraph this fragment belongs to (widow/orphan control).</summary>
        public object ParagraphId;
        /// <summary>Index of this line within its paragraph, and the paragraph's line count.</summary>
        public int LineIndex;
        public int LineCount;
        /// <summary>False when the paragraph turned widow/orphan control off.</summary>
        public bool WidowControl = true;
        /// <summary>
        /// Horizontal border widths included in this row fragment's height. Word counts
        /// them in a header's occupied height but not in a footer's — the footer anchors
        /// at the page bottom and its top border overlaps the body area (PROBE13/14 via
        /// COM; SaBC p26: the embedded map fits right under the border-free footer limit).
        /// </summary>
        public double EdgeExtentPt;

        public Fragment() { }
        public Fragment(double height) { Height = height; }

        public void Add(DrawOp op)
        {
            if (op != null)
                Ops.Add(op);
        }

        public void AddTranslated(IEnumerable<DrawOp> ops, double dx, double dy)
        {
            foreach (var op in ops)
                Ops.Add(op.Translate(dx, dy));
        }
    }

    internal sealed class LayoutPage
    {
        public double Width = 612;
        public double Height = 792;
        public List<DrawOp> Ops = new List<DrawOp>();
        public int Number;              // 1-based page number as shown in fields
    }

    internal sealed class OutlineEntry
    {
        public int Level;
        public string Title;
        public int PageIndex;
        public double Y;
        public readonly List<OutlineEntry> Children = new List<OutlineEntry>();
    }

    internal sealed class LayoutResult
    {
        public readonly List<LayoutPage> Pages = new List<LayoutPage>();
        public readonly List<OutlineEntry> Outline = new List<OutlineEntry>();
        public readonly Dictionary<string, KeyValuePair<int, double>> Bookmarks =
            new Dictionary<string, KeyValuePair<int, double>>(StringComparer.OrdinalIgnoreCase);
    }
}
