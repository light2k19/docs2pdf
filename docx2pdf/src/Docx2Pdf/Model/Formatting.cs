using System.Collections.Generic;

namespace Docx2Pdf.Model
{
    internal enum TextAlignment { Left, Center, Right, Justify }

    internal enum LineSpacingRule { Auto, AtLeast, Exact }

    internal enum VerticalTextAlignment { Baseline, Superscript, Subscript }

    internal enum UnderlineStyle { None, Single, Double, Thick, Dotted, Dashed, Wave }

    internal enum BorderStyle { None, Single, Double, Dotted, Dashed, Thick }

    internal enum VerticalCellAlignment { Top, Center, Bottom }

    internal enum TabAlignment { Left, Center, Right, Decimal, Bar, Clear }

    internal enum TabLeader { None, Dot, Hyphen, Underscore }

    internal sealed class TabStop
    {
        public double PositionPt;
        public TabAlignment Alignment = TabAlignment.Left;
        public TabLeader Leader = TabLeader.None;
    }

    internal sealed class Border
    {
        public BorderStyle Style = BorderStyle.None;
        public double WidthPt = 0.5;
        public uint Color = 0x000000;
        public double SpacePt;

        public bool IsVisible { get { return Style != BorderStyle.None && WidthPt > 0; } }

        public Border Clone()
        {
            return new Border { Style = Style, WidthPt = WidthPt, Color = Color, SpacePt = SpacePt };
        }
    }

    internal sealed class Borders
    {
        public Border Top, Bottom, Left, Right, InsideH, InsideV;

        public Borders Clone()
        {
            return new Borders
            {
                Top = Top == null ? null : Top.Clone(),
                Bottom = Bottom == null ? null : Bottom.Clone(),
                Left = Left == null ? null : Left.Clone(),
                Right = Right == null ? null : Right.Clone(),
                InsideH = InsideH == null ? null : InsideH.Clone(),
                InsideV = InsideV == null ? null : InsideV.Clone(),
            };
        }

        public bool Any
        {
            get
            {
                return (Top != null && Top.IsVisible) || (Bottom != null && Bottom.IsVisible)
                    || (Left != null && Left.IsVisible) || (Right != null && Right.IsVisible);
            }
        }
    }

    /// <summary>Character (run) level formatting. Null members mean "inherit".</summary>
    internal sealed class CharacterFormat
    {
        public string FontFamily;
        public string EastAsianFontFamily;
        public bool? Bold;
        public bool? Italic;
        public UnderlineStyle? Underline;
        public uint? UnderlineColor;
        public bool? Strike;
        public double? SizePt;
        public uint? Color;
        public uint? Highlight;
        public uint? Shading;
        public VerticalTextAlignment? VertAlign;
        public bool? AllCaps;
        public bool? SmallCaps;
        public bool? Hidden;
        /// <summary>Extra inter-character spacing, in points.</summary>
        public double? CharacterSpacingPt;
        /// <summary>w:position — vertical run offset in points; positive raises, negative lowers.</summary>
        public double? RaisePt;

        public CharacterFormat Clone()
        {
            return (CharacterFormat)MemberwiseClone();
        }

        /// <summary>Applies non-null values of <paramref name="other"/> on top of this instance.</summary>
        public void ApplyOver(CharacterFormat other)
        {
            if (other == null)
                return;
            if (other.FontFamily != null) FontFamily = other.FontFamily;
            if (other.EastAsianFontFamily != null) EastAsianFontFamily = other.EastAsianFontFamily;
            if (other.Bold.HasValue) Bold = other.Bold;
            if (other.Italic.HasValue) Italic = other.Italic;
            if (other.Underline.HasValue) Underline = other.Underline;
            if (other.UnderlineColor.HasValue) UnderlineColor = other.UnderlineColor;
            if (other.Strike.HasValue) Strike = other.Strike;
            if (other.SizePt.HasValue) SizePt = other.SizePt;
            if (other.Color.HasValue) Color = other.Color;
            if (other.Highlight.HasValue) Highlight = other.Highlight;
            if (other.Shading.HasValue) Shading = other.Shading;
            if (other.VertAlign.HasValue) VertAlign = other.VertAlign;
            if (other.AllCaps.HasValue) AllCaps = other.AllCaps;
            if (other.SmallCaps.HasValue) SmallCaps = other.SmallCaps;
            if (other.Hidden.HasValue) Hidden = other.Hidden;
            if (other.RaisePt.HasValue) RaisePt = other.RaisePt;
            if (other.CharacterSpacingPt.HasValue) CharacterSpacingPt = other.CharacterSpacingPt;
        }

        public static CharacterFormat Default()
        {
            return new CharacterFormat
            {
                FontFamily = "Calibri",
                Bold = false,
                Italic = false,
                Underline = UnderlineStyle.None,
                Strike = false,
                SizePt = 11,
                Color = 0x000000,
                VertAlign = VerticalTextAlignment.Baseline,
                AllCaps = false,
                SmallCaps = false,
                Hidden = false,
            };
        }
    }

    /// <summary>Paragraph level formatting. Null members mean "inherit".</summary>
    internal sealed class ParagraphFormat
    {
        public TextAlignment? Alignment;
        public double? IndentLeftPt;
        public double? IndentRightPt;
        /// <summary>Positive = first line indent, negative = hanging indent.</summary>
        public double? IndentFirstLinePt;
        public double? SpaceBeforePt;
        public double? SpaceAfterPt;
        public bool? ContextualSpacing;
        public double? LineSpacing;             // multiple (Auto) or points (AtLeast/Exact)
        public LineSpacingRule? LineSpacingRule;
        public bool? KeepNext;
        public bool? KeepLines;
        public bool? PageBreakBefore;
        public Borders Borders;
        public uint? Shading;
        public List<TabStop> Tabs;
        public int? OutlineLevel;
        public bool? Bidi;
        /// <summary>Word's widow/orphan control; on by default.</summary>
        public bool? WidowControl;
        /// <summary>HTML "auto" spacing flags (w:beforeAutospacing / w:afterAutospacing).</summary>
        public bool? AutoSpaceBefore;
        public bool? AutoSpaceAfter;
        /// <summary>w:framePr w:dropCap — this paragraph is a drop cap spanning that many lines.</summary>
        public int? DropCapLines;
        /// <summary>True when SpaceBeforePt came from the paragraph's own pPr rather than a style.
        /// Set by the reader on resolved formats; not merged by ApplyOver.</summary>
        public bool SpaceBeforeIsDirect;
        /// <summary>True when SpaceAfterPt came from the paragraph's own pPr rather than a style.</summary>
        public bool SpaceAfterIsDirect;

        public ParagraphFormat Clone()
        {
            var c = (ParagraphFormat)MemberwiseClone();
            c.Borders = Borders == null ? null : Borders.Clone();
            c.Tabs = Tabs == null ? null : new List<TabStop>(Tabs);
            return c;
        }

        public void ApplyOver(ParagraphFormat other)
        {
            if (other == null)
                return;
            if (other.Alignment.HasValue) Alignment = other.Alignment;
            if (other.IndentLeftPt.HasValue) IndentLeftPt = other.IndentLeftPt;
            if (other.IndentRightPt.HasValue) IndentRightPt = other.IndentRightPt;
            if (other.IndentFirstLinePt.HasValue) IndentFirstLinePt = other.IndentFirstLinePt;
            if (other.SpaceBeforePt.HasValue) SpaceBeforePt = other.SpaceBeforePt;
            if (other.SpaceAfterPt.HasValue) SpaceAfterPt = other.SpaceAfterPt;
            if (other.ContextualSpacing.HasValue) ContextualSpacing = other.ContextualSpacing;
            if (other.LineSpacing.HasValue) { LineSpacing = other.LineSpacing; LineSpacingRule = other.LineSpacingRule; }
            if (other.KeepNext.HasValue) KeepNext = other.KeepNext;
            if (other.KeepLines.HasValue) KeepLines = other.KeepLines;
            if (other.PageBreakBefore.HasValue) PageBreakBefore = other.PageBreakBefore;
            if (other.Shading.HasValue) Shading = other.Shading;
            if (other.OutlineLevel.HasValue) OutlineLevel = other.OutlineLevel;
            if (other.Bidi.HasValue) Bidi = other.Bidi;
            if (other.WidowControl.HasValue) WidowControl = other.WidowControl;
            if (other.AutoSpaceBefore.HasValue) AutoSpaceBefore = other.AutoSpaceBefore;
            if (other.AutoSpaceAfter.HasValue) AutoSpaceAfter = other.AutoSpaceAfter;
            if (other.DropCapLines.HasValue) DropCapLines = other.DropCapLines;
            if (other.Tabs != null && other.Tabs.Count > 0)
            {
                // A cleared stop removes the inherited stop at that position; the clear
                // entry itself is not a stop (the licence form's footer clears the Footer
                // style's centre/right stops before adding its own page-number tab).
                var list = Tabs == null ? new List<TabStop>() : new List<TabStop>(Tabs);
                foreach (var stop in other.Tabs)
                {
                    if (stop.Alignment == TabAlignment.Clear)
                        list.RemoveAll(t => System.Math.Abs(t.PositionPt - stop.PositionPt) < 0.05);
                    else
                        list.Add(stop);
                }
                Tabs = list;
            }
            if (other.Borders != null)
            {
                if (Borders == null)
                {
                    Borders = other.Borders.Clone();
                }
                else
                {
                    if (other.Borders.Top != null) Borders.Top = other.Borders.Top.Clone();
                    if (other.Borders.Bottom != null) Borders.Bottom = other.Borders.Bottom.Clone();
                    if (other.Borders.Left != null) Borders.Left = other.Borders.Left.Clone();
                    if (other.Borders.Right != null) Borders.Right = other.Borders.Right.Clone();
                }
            }
        }

        public static ParagraphFormat Default()
        {
            return new ParagraphFormat
            {
                Alignment = TextAlignment.Left,
                IndentLeftPt = 0,
                IndentRightPt = 0,
                IndentFirstLinePt = 0,
                SpaceBeforePt = 0,
                SpaceAfterPt = 0,
                ContextualSpacing = false,
                LineSpacing = 1.0,
                LineSpacingRule = Model.LineSpacingRule.Auto,
                KeepNext = false,
                KeepLines = false,
                PageBreakBefore = false,
                OutlineLevel = null,
                Bidi = false,
            };
        }
    }
}
