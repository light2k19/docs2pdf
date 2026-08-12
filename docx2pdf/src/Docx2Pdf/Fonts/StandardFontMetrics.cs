using System;
using System.Collections.Generic;

namespace Docx2Pdf.Fonts
{
    internal enum StandardFontFamily { Helvetica, Times, Courier, Symbol }

    /// <summary>
    /// Adobe base-14 font metrics (widths in 1/1000 em, WinAnsi code indexed).
    /// Used only when no system font file can be embedded.
    /// </summary>
    internal static class StandardFontMetrics
    {
        // ASCII 32..126 widths.
        private static readonly short[] HelveticaAscii =
        {
            278,278,355,556,556,889,667,222,333,333,389,584,278,333,278,278,
            556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,
            1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,
            667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,
            222,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,
            556,556,333,500,278,556,500,722,500,500,500,334,260,334,584
        };

        private static readonly short[] HelveticaBoldAscii =
        {
            278,333,474,556,556,889,722,278,333,333,389,584,278,333,278,278,
            556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,
            975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,
            667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,
            278,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,
            611,611,389,556,333,611,556,778,556,556,500,389,280,389,584
        };

        private static readonly short[] TimesAscii =
        {
            250,333,408,500,500,833,778,333,333,333,500,564,250,333,250,278,
            500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,
            921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,
            556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,
            333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,
            500,500,333,389,278,500,500,722,500,500,444,480,200,480,541
        };

        private static readonly short[] TimesBoldAscii =
        {
            250,333,555,500,500,1000,833,333,333,333,500,570,250,333,250,278,
            500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,
            930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,
            611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,
            333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,
            556,556,444,389,333,556,500,722,500,500,444,394,220,394,520
        };

        private static readonly short[] TimesItalicAscii =
        {
            250,333,420,500,500,833,778,333,333,333,500,675,250,333,250,278,
            500,500,500,500,500,500,500,500,500,500,333,333,675,675,675,500,
            920,611,611,667,722,611,611,722,722,333,444,667,556,833,667,722,
            611,722,611,500,556,722,611,833,611,556,556,389,278,389,422,500,
            333,500,500,444,500,444,278,500,500,278,278,444,278,722,500,500,
            500,500,389,389,278,500,444,667,444,444,389,400,275,400,541
        };

        private static readonly short[] TimesBoldItalicAscii =
        {
            250,389,555,500,500,833,778,333,333,333,500,570,250,333,250,278,
            500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,
            832,667,667,667,722,667,667,722,778,389,500,667,611,889,722,722,
            611,722,667,556,611,722,667,889,667,611,611,333,278,333,570,500,
            333,500,500,444,500,444,333,500,556,278,278,500,278,778,556,500,
            500,500,389,389,278,556,444,667,500,444,389,348,220,348,570
        };

        // High-range (0x80-0xFF) widths for Helvetica and Times; bold/italic reuse these.
        private static readonly short[] HelveticaHigh = BuildHigh(
            HelveticaAscii,
            new Dictionary<int, short>
            {
                {0x80,556},{0x82,222},{0x83,556},{0x84,333},{0x85,1000},{0x86,556},{0x87,556},{0x88,333},
                {0x89,1000},{0x8A,667},{0x8B,333},{0x8C,1000},{0x8E,611},{0x91,222},{0x92,222},{0x93,333},
                {0x94,333},{0x95,350},{0x96,556},{0x97,1000},{0x98,333},{0x99,1000},{0x9A,500},{0x9B,333},
                {0x9C,944},{0x9E,500},{0x9F,667},{0xA0,278},{0xA1,333},{0xA2,556},{0xA3,556},{0xA4,556},
                {0xA5,556},{0xA6,260},{0xA7,556},{0xA8,333},{0xA9,737},{0xAA,370},{0xAB,556},{0xAC,584},
                {0xAD,333},{0xAE,737},{0xAF,333},{0xB0,400},{0xB1,584},{0xB2,333},{0xB3,333},{0xB4,333},
                {0xB5,556},{0xB6,537},{0xB7,278},{0xB8,333},{0xB9,333},{0xBA,365},{0xBB,556},{0xBC,834},
                {0xBD,834},{0xBE,834},{0xBF,611},{0xC6,1000},{0xD0,722},{0xD7,584},{0xD8,778},{0xDE,667},
                {0xDF,611},{0xE6,889},{0xF0,556},{0xF7,584},{0xF8,611},{0xFE,556},
            });

        private static readonly short[] TimesHigh = BuildHigh(
            TimesAscii,
            new Dictionary<int, short>
            {
                {0x80,500},{0x82,333},{0x83,500},{0x84,444},{0x85,1000},{0x86,500},{0x87,500},{0x88,333},
                {0x89,1000},{0x8A,556},{0x8B,333},{0x8C,889},{0x8E,611},{0x91,333},{0x92,333},{0x93,444},
                {0x94,444},{0x95,350},{0x96,500},{0x97,1000},{0x98,333},{0x99,980},{0x9A,389},{0x9B,333},
                {0x9C,722},{0x9E,444},{0x9F,722},{0xA0,250},{0xA1,333},{0xA2,500},{0xA3,500},{0xA4,500},
                {0xA5,500},{0xA6,200},{0xA7,500},{0xA8,333},{0xA9,760},{0xAA,276},{0xAB,500},{0xAC,564},
                {0xAD,333},{0xAE,760},{0xAF,333},{0xB0,400},{0xB1,564},{0xB2,300},{0xB3,300},{0xB4,333},
                {0xB5,500},{0xB6,453},{0xB7,250},{0xB8,333},{0xB9,300},{0xBA,310},{0xBB,500},{0xBC,750},
                {0xBD,750},{0xBE,750},{0xBF,500},{0xC6,889},{0xD0,722},{0xD7,564},{0xD8,722},{0xDE,556},
                {0xDF,500},{0xE6,667},{0xF0,500},{0xF7,564},{0xF8,500},{0xFE,500},
            });

        /// <summary>
        /// Builds a 256 entry width table: accented letters inherit the width of their base letter,
        /// everything else comes from the explicit symbol table.
        /// </summary>
        private static short[] BuildHigh(short[] ascii, Dictionary<int, short> symbols)
        {
            var widths = new short[256];
            for (var i = 32; i < 127; i++)
                widths[i] = ascii[i - 32];

            // WinAnsi 0xC0-0xFF are accented Latin letters; borrow the base letter width.
            const string bases = "AAAAAAECEEEEIIIIDNOOOOOxOUUUUYPsaaaaaaeceeeeiiiidnooooo/ouuuuypy";
            for (var i = 0xC0; i <= 0xFF; i++)
            {
                var b = bases[i - 0xC0];
                if (b != 'x' && b != '/')
                    widths[i] = widths[b];
            }

            foreach (var kv in symbols)
                widths[kv.Key] = kv.Value;

            for (var i = 0; i < widths.Length; i++)
            {
                if (widths[i] == 0)
                    widths[i] = ascii[0];   // fall back to the space width
            }
            return widths;
        }

        public static short[] GetWidths(StandardFontFamily family, bool bold, bool italic)
        {
            switch (family)
            {
                case StandardFontFamily.Courier:
                {
                    var w = new short[256];
                    for (var i = 0; i < 256; i++)
                        w[i] = 600;
                    return w;
                }
                case StandardFontFamily.Times:
                {
                    var ascii = bold && italic ? TimesBoldItalicAscii
                              : bold ? TimesBoldAscii
                              : italic ? TimesItalicAscii
                              : TimesAscii;
                    return Merge(ascii, TimesHigh);
                }
                default:
                {
                    var ascii = bold ? HelveticaBoldAscii : HelveticaAscii;
                    return Merge(ascii, HelveticaHigh);
                }
            }
        }

        private static short[] Merge(short[] ascii, short[] high)
        {
            var widths = new short[256];
            Array.Copy(high, widths, 256);
            for (var i = 32; i < 127; i++)
                widths[i] = ascii[i - 32];
            return widths;
        }

        public static string BaseFontName(StandardFontFamily family, bool bold, bool italic)
        {
            switch (family)
            {
                case StandardFontFamily.Courier:
                    if (bold && italic) return "Courier-BoldOblique";
                    if (bold) return "Courier-Bold";
                    if (italic) return "Courier-Oblique";
                    return "Courier";
                case StandardFontFamily.Times:
                    if (bold && italic) return "Times-BoldItalic";
                    if (bold) return "Times-Bold";
                    if (italic) return "Times-Italic";
                    return "Times-Roman";
                default:
                    if (bold && italic) return "Helvetica-BoldOblique";
                    if (bold) return "Helvetica-Bold";
                    if (italic) return "Helvetica-Oblique";
                    return "Helvetica";
            }
        }

        /// <summary>Picks the closest base-14 family for an arbitrary font name.</summary>
        public static StandardFontFamily ClassifyFamily(string name)
        {
            if (string.IsNullOrEmpty(name))
                return StandardFontFamily.Helvetica;
            var n = name.ToLowerInvariant();
            if (n.Contains("courier") || n.Contains("mono") || n.Contains("consol"))
                return StandardFontFamily.Courier;
            if (n.Contains("times") || n.Contains("serif") || n.Contains("georgia") || n.Contains("garamond")
                || n.Contains("book") || n.Contains("roman") || n.Contains("cambria") || n.Contains("palatino")
                || n.Contains("minion") || n.Contains("constantia"))
            {
                // "sans serif" must not be treated as a serif face.
                if (!n.Contains("sans"))
                    return StandardFontFamily.Times;
            }
            return StandardFontFamily.Helvetica;
        }
    }
}
