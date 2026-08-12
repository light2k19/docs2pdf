using System.Collections.Generic;

namespace Docx2Pdf.Fonts
{
    /// <summary>Unicode &lt;-&gt; WinAnsi (CP1252) mapping used by the base-14 fallback fonts.</summary>
    internal static class WinAnsiEncoding
    {
        private static readonly Dictionary<int, byte> Map = BuildMap();
        private static readonly int[] Reverse = BuildReverse();

        private static Dictionary<int, byte> BuildMap()
        {
            var map = new Dictionary<int, byte>(256);
            for (var i = 32; i < 127; i++)
                map[i] = (byte)i;
            for (var i = 0xA0; i <= 0xFF; i++)
                map[i] = (byte)i;

            // CP1252 specials in the 0x80-0x9F range.
            map[0x20AC] = 0x80; map[0x201A] = 0x82; map[0x0192] = 0x83; map[0x201E] = 0x84;
            map[0x2026] = 0x85; map[0x2020] = 0x86; map[0x2021] = 0x87; map[0x02C6] = 0x88;
            map[0x2030] = 0x89; map[0x0160] = 0x8A; map[0x2039] = 0x8B; map[0x0152] = 0x8C;
            map[0x017D] = 0x8E; map[0x2018] = 0x91; map[0x2019] = 0x92; map[0x201C] = 0x93;
            map[0x201D] = 0x94; map[0x2022] = 0x95; map[0x2013] = 0x96; map[0x2014] = 0x97;
            map[0x02DC] = 0x98; map[0x2122] = 0x99; map[0x0161] = 0x9A; map[0x203A] = 0x9B;
            map[0x0153] = 0x9C; map[0x017E] = 0x9E; map[0x0178] = 0x9F;

            // Convenient approximations so common typography still renders.
            map[0x00A0] = 0xA0;    // no-break space
            map[0x2011] = 0x2D;    // non-breaking hyphen -> hyphen
            map[0x2012] = 0x96;
            map[0x2212] = 0x2D;
            map[0x2010] = 0x2D;
            map[0x00AD] = 0x2D;
            map[0x2027] = 0xB7;
            map[0x25CF] = 0x95;
            map[0x25AA] = 0x95;
            map[0x25A0] = 0x95;
            map[0x25E6] = 0x6F;
            map[0x2043] = 0x2D;
            map[0x00B7] = 0xB7;
            map[0x2192] = 0x3E;    // arrows -> angle brackets
            map[0x2190] = 0x3C;
            map[0x2009] = 0x20;
            map[0x200B] = 0x20;
            map[0x2003] = 0x20;
            map[0x2002] = 0x20;
            map[0x0009] = 0x20;
            return map;
        }

        private static int[] BuildReverse()
        {
            var rev = new int[256];
            foreach (var kv in Map)
            {
                // Prefer the canonical code point for each byte.
                if (rev[kv.Value] == 0 || kv.Key < 0x100)
                    rev[kv.Value] = kv.Key;
            }
            return rev;
        }

        public static bool TryGetByte(int codePoint, out byte value)
        {
            return Map.TryGetValue(codePoint, out value);
        }

        /// <summary>Maps a byte back to Unicode; used when building ToUnicode CMaps.</summary>
        public static int ToUnicode(byte code)
        {
            var u = Reverse[code];
            return u == 0 ? code : u;
        }
    }
}
