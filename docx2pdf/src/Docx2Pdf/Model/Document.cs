using System.Collections.Generic;

namespace Docx2Pdf.Model
{
    internal abstract class Block
    {
    }

    internal enum BreakKind { Line, Page, Column }

    internal enum FieldKind { None, Page, NumPages, SectionPages }

    internal abstract class Inline
    {
        public CharacterFormat Format;
        /// <summary>External URL, when this inline sits inside a hyperlink.</summary>
        public string LinkUrl;
        /// <summary>Internal bookmark name, when this inline links inside the document.</summary>
        public string LinkAnchor;
    }

    internal sealed class TextInline : Inline
    {
        public string Text;
        public TextInline() { }
        public TextInline(string text, CharacterFormat fmt) { Text = text; Format = fmt; }
    }

    internal sealed class BreakInline : Inline
    {
        public BreakKind Kind = BreakKind.Line;
    }

    internal sealed class TabInline : Inline
    {
    }

    internal sealed class FieldInline : Inline
    {
        public FieldKind Kind;
    }

    internal sealed class ImageInline : Inline
    {
        /// <summary>Package part name of the image, used as the cache key.</summary>
        public string PartName;
        public byte[] Data;
        public double WidthPt;
        public double HeightPt;
        /// <summary>Clockwise rotation in degrees (from a:xfrm/@rot).</summary>
        public double RotationDeg;
        public string Description;
        /// <summary>Extra layout space around the picture for effects such as shadows (wp:effectExtent).</summary>
        public double EffectLeftPt, EffectTopPt, EffectRightPt, EffectBottomPt;
        /// <summary>Picture frame stroke width (a:ln); the stroke straddles the extent edge.</summary>
        public double FrameWidthPt;
        /// <summary>Frame stroke colour as RRGGBB hex, null when unset.</summary>
        public string FrameColor;
        /// <summary>Source crop fractions 0..1 per edge (a:srcRect, thousandths of a percent).</summary>
        public double CropLeft, CropTop, CropRight, CropBottom;
        /// <summary>
        /// Annotation shapes a group (wpg:wgp) draws over the picture — highlight
        /// rectangles and the like — in points relative to the picture box's top-left.
        /// </summary>
        public System.Collections.Generic.List<ImageOverlay> Overlays;
    }

    /// <summary>A simple rectangle a group shape draws over its picture.</summary>
    internal sealed class ImageOverlay
    {
        public double X, Y, Width, Height;
        public uint? OutlineColor;
        public double OutlineWidthPt;
        public uint? FillColor;
    }

    internal sealed class BookmarkInline : Inline
    {
        public string Name;
    }

    /// <summary>What a floating object's position is measured from.</summary>
    internal enum AnchorRelativeFrom
    {
        Page, Margin, LeftMargin, RightMargin, TopMargin, BottomMargin,
        Column, Character, Paragraph, Line, InsideMargin, OutsideMargin
    }

    /// <summary>
    /// A floating (anchored) drawing: a picture or a text box positioned relative to the page,
    /// the margins or the anchoring paragraph rather than flowing with the text.
    /// </summary>
    internal sealed class AnchoredInline : Inline
    {
        /// <summary>Content of the frame; a picture is wrapped in a single paragraph.</summary>
        public List<Block> Blocks = new List<Block>();
        public double WidthPt;
        public double HeightPt;

        public AnchorRelativeFrom HorizontalFrom = AnchorRelativeFrom.Column;
        public double HorizontalOffsetPt;
        /// <summary>left / center / right / inside / outside; null when an offset is used.</summary>
        public string HorizontalAlign;

        public AnchorRelativeFrom VerticalFrom = AnchorRelativeFrom.Paragraph;
        public double VerticalOffsetPt;
        /// <summary>top / center / bottom / inside / outside; null when an offset is used.</summary>
        public string VerticalAlign;

        /// <summary>Drawn behind the text when set.</summary>
        public bool BehindDoc;

        /// <summary>Simple preset-geometry shape styling (highlight boxes and the like).</summary>
        public uint? OutlineColor;
        public double OutlineWidthPt = 1;
        public uint? FillColor;
    }

    internal sealed class Paragraph : Block
    {
        public ParagraphFormat Format = new ParagraphFormat();
        public CharacterFormat RunDefaults = new CharacterFormat();
        /// <summary>
        /// Formatting of the paragraph mark (w:pPr/w:rPr). It applies to the pilcrow only —
        /// it sizes empty paragraphs but does not restyle the paragraph's text runs.
        /// </summary>
        public CharacterFormat MarkFormat;
        public List<Inline> Inlines = new List<Inline>();
        public string StyleId;
        /// <summary>Resolved list label ("1.", "a)", "&#x2022;"), or null when not numbered.</summary>
        public string ListLabel;
        /// <summary>Numbering instance (w:numId) the paragraph belongs to, null when not numbered.</summary>
        public string ListNumId;
        public CharacterFormat ListLabelFormat;
        /// <summary>Text position (points) the paragraph body starts at when numbered.</summary>
        public double ListTextIndentPt;
        public double ListLabelIndentPt;
        public bool ListLabelFollowedByTab = true;
        /// <summary>Heading level 1..9 for PDF outline generation, 0 when not a heading.</summary>
        public int HeadingLevel;
    }

    internal enum VerticalMerge { None, Restart, Continue }

    internal sealed class TableCell
    {
        public List<Block> Blocks = new List<Block>();
        public double WidthPt;                 // preferred width, 0 = auto
        public bool WidthIsPercent;
        public int GridSpan = 1;
        public VerticalMerge VMerge = VerticalMerge.None;
        public Borders Borders;
        public uint? Shading;
        public VerticalCellAlignment VAlign = VerticalCellAlignment.Top;
        public double? MarginLeftPt, MarginRightPt, MarginTopPt, MarginBottomPt;
        /// <summary>Text direction btLr / tbRl produce rotated cells; only detected, not rotated.</summary>
        public bool Vertical;
    }

    internal sealed class TableRow
    {
        public List<TableCell> Cells = new List<TableCell>();
        public double HeightPt;
        public bool HeightExact;
        public bool IsHeader;
        public bool CantSplit;
    }

    internal sealed class Table : Block
    {
        public List<TableRow> Rows = new List<TableRow>();
        /// <summary>Column widths in points, from w:tblGrid.</summary>
        public List<double> Grid = new List<double>();
        public Borders Borders;
        public uint? Shading;
        public double? PreferredWidthPt;
        public bool PreferredWidthIsPercent;
        public TextAlignment Alignment = TextAlignment.Left;
        public double IndentPt;
        public double CellMarginLeftPt = 5.4;
        public double CellMarginRightPt = 5.4;
        public double CellMarginTopPt;
        public double CellMarginBottomPt;
        public double CellSpacingPt;
    }

    internal sealed class SectionProperties
    {
        public double PageWidthPt = 612;
        public double PageHeightPt = 792;
        public double MarginLeftPt = 72;
        public double MarginRightPt = 72;
        public double MarginTopPt = 72;
        public double MarginBottomPt = 72;
        public double HeaderDistancePt = 36;
        public double FooterDistancePt = 36;
        public double GutterPt;
        public bool Landscape;
        public bool TitlePage;
        public int Columns = 1;
        public double ColumnSpacingPt = 36;
        public int? PageNumberStart;
        /// <summary>w:pgNumType/@w:fmt, e.g. "lowerRoman"; null means decimal.</summary>
        public string PageNumberFormat;

        public HeaderFooter HeaderDefault, HeaderFirst, HeaderEven;
        public HeaderFooter FooterDefault, FooterFirst, FooterEven;

        /// <summary>Page borders drawn around the whole page (w:pgBorders).</summary>
        public Borders PageBorders;
        /// <summary>True when the border offsets are measured from the text margins rather than the page edge.</summary>
        public bool PageBordersFromText;
        /// <summary>w:pgBorders/@w:display — allPages, firstPage or notFirstPage.</summary>
        public string PageBordersDisplay;

        public double ContentWidthPt { get { return PageWidthPt - MarginLeftPt - MarginRightPt; } }
        public double ContentHeightPt { get { return PageHeightPt - MarginTopPt - MarginBottomPt; } }

        public SectionProperties Clone()
        {
            return (SectionProperties)MemberwiseClone();
        }
    }

    internal sealed class HeaderFooter
    {
        public List<Block> Blocks = new List<Block>();
    }

    internal sealed class Section
    {
        /// <summary>Null until a sectPr is seen; the reader fills in the document defaults afterwards.</summary>
        public SectionProperties Properties;
        public List<Block> Blocks = new List<Block>();
        /// <summary>True when the section starts on a new page (w:type != continuous).</summary>
        public bool StartsNewPage = true;
    }

    internal sealed class DocumentInfo
    {
        public string Title;
        public string Author;
        public string Subject;
        public string Keywords;
        public string Creator;
    }

    internal sealed class WordDocument
    {
        public List<Section> Sections = new List<Section>();
        public DocumentInfo Info = new DocumentInfo();
        public bool EvenOddHeaders;
        public double DefaultTabStopPt = 36;
        /// <summary>Footnote/endnote bodies keyed by id, rendered at the end of the document.</summary>
        public List<KeyValuePair<string, List<Block>>> Notes = new List<KeyValuePair<string, List<Block>>>();
    }
}
