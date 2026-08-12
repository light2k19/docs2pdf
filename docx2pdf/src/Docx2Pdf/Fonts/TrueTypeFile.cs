using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Docx2Pdf.Fonts
{
    /// <summary>
    /// Minimal SFNT (TrueType / OpenType / TrueType collection) reader: enough to measure text,
    /// map Unicode to glyph ids and embed the font program in a PDF.
    /// </summary>
    internal sealed class TrueTypeFile
    {
        private sealed class TableRecord
        {
            public uint Offset;
            public uint Length;
        }

        private readonly Dictionary<string, TableRecord> _tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        private byte[] _data;
        private ushort[] _advanceWidths;
        private Dictionary<int, int> _cmap;
        private Dictionary<int, int> _symbolCmap;

        public string FilePath;
        public int CollectionIndex;
        public bool IsCff;                 // OpenType/CFF outlines (embedded as FontFile3)
        public string PostScriptName;
        public string FamilyName;
        public string SubfamilyName;
        public string TypographicFamily;
        public int UnitsPerEm = 1000;
        public int NumGlyphs;
        public short XMin, YMin, XMax, YMax;
        public short Ascender, Descender, LineGap;
        public short TypoAscender, TypoDescender, TypoLineGap;
        /// <summary>OS/2 fsSelection bit 7 — the font asks layout to use the sTypo metrics.</summary>
        public bool UseTypoMetrics;
        /// <summary>OS/2 usWinAscent / usWinDescent — what Word/GDI use to place the baseline.</summary>
        public ushort WinAscent, WinDescent;
        public short CapHeight;
        public ushort WeightClass = 400;
        public bool IsBold, IsItalic, IsFixedPitch, IsSerif;
        public double ItalicAngle;

        public byte[] Data { get { return _data; } }

        /// <summary>Loads the whole font program (rebuilding a standalone SFNT for collections).</summary>
        public static TrueTypeFile Load(string path, int collectionIndex)
        {
            var bytes = File.ReadAllBytes(path);
            var font = new TrueTypeFile { FilePath = path, CollectionIndex = collectionIndex };
            font.Initialize(bytes, collectionIndex);
            return font;
        }

        /// <summary>Reads only the metadata tables needed to index a font file.</summary>
        public static List<TrueTypeFile> Probe(string path)
        {
            var result = new List<TrueTypeFile>();
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096))
            {
                var header = ReadBytes(stream, 0, 12);
                if (header == null)
                    return result;

                var tag = Encoding.ASCII.GetString(header, 0, 4);
                var offsets = new List<uint>();
                if (tag == "ttcf")
                {
                    var count = (int)ReadU32(header, 8);
                    if (count <= 0 || count > 1024)
                        return result;
                    var table = ReadBytes(stream, 12, count * 4);
                    if (table == null)
                        return result;
                    for (var i = 0; i < count; i++)
                        offsets.Add(ReadU32(table, i * 4));
                }
                else
                {
                    offsets.Add(0);
                }

                for (var i = 0; i < offsets.Count; i++)
                {
                    var font = ProbeOne(stream, offsets[i]);
                    if (font == null)
                        continue;
                    font.FilePath = path;
                    font.CollectionIndex = i;
                    result.Add(font);
                }
            }
            return result;
        }

        private static TrueTypeFile ProbeOne(Stream stream, uint offset)
        {
            var dir = ReadBytes(stream, offset, 12);
            if (dir == null)
                return null;
            var version = ReadU32(dir, 0);
            if (version != 0x00010000 && version != 0x74727565 && version != 0x4F54544F)
                return null;
            var numTables = ReadU16(dir, 4);
            if (numTables == 0 || numTables > 512)
                return null;

            var records = ReadBytes(stream, offset + 12, numTables * 16);
            if (records == null)
                return null;

            var font = new TrueTypeFile { IsCff = version == 0x4F54544F };
            for (var i = 0; i < numTables; i++)
            {
                var tag = Encoding.ASCII.GetString(records, i * 16, 4);
                font._tables[tag] = new TableRecord
                {
                    Offset = ReadU32(records, i * 16 + 8),
                    Length = ReadU32(records, i * 16 + 12),
                };
            }

            font.ReadHeadFromStream(stream);
            font.ReadOs2FromStream(stream);
            font.ReadNamesFromStream(stream);
            return font;
        }

        private static byte[] ReadBytes(Stream stream, long offset, int count)
        {
            if (offset < 0 || count <= 0 || offset + count > stream.Length)
                return null;
            stream.Position = offset;
            var buffer = new byte[count];
            var read = 0;
            while (read < count)
            {
                var n = stream.Read(buffer, read, count - read);
                if (n <= 0)
                    return null;
                read += n;
            }
            return buffer;
        }

        private void ReadHeadFromStream(Stream stream)
        {
            TableRecord head;
            if (!_tables.TryGetValue("head", out head))
                return;
            var bytes = ReadBytes(stream, head.Offset, (int)Math.Min(head.Length, 54u));
            if (bytes == null || bytes.Length < 54)
                return;
            UnitsPerEm = ReadU16(bytes, 18);
            if (UnitsPerEm == 0)
                UnitsPerEm = 1000;
            XMin = (short)ReadU16(bytes, 36);
            YMin = (short)ReadU16(bytes, 38);
            XMax = (short)ReadU16(bytes, 40);
            YMax = (short)ReadU16(bytes, 42);
            var macStyle = ReadU16(bytes, 44);
            IsBold = (macStyle & 1) != 0;
            IsItalic = (macStyle & 2) != 0;
        }

        private void ReadOs2FromStream(Stream stream)
        {
            TableRecord os2;
            if (!_tables.TryGetValue("OS/2", out os2))
                return;
            var bytes = ReadBytes(stream, os2.Offset, (int)Math.Min(os2.Length, 96u));
            if (bytes == null || bytes.Length < 64)
                return;
            WeightClass = ReadU16(bytes, 4);
            var familyClass = (short)ReadU16(bytes, 30);
            IsSerif = (familyClass >> 8) >= 1 && (familyClass >> 8) <= 7;
            var fsSelection = ReadU16(bytes, 62);
            if ((fsSelection & 0x20) != 0) IsBold = true;
            if ((fsSelection & 0x01) != 0) IsItalic = true;
            UseTypoMetrics = (fsSelection & 0x80) != 0;
            if (bytes.Length >= 74)
            {
                TypoAscender = (short)ReadU16(bytes, 68);
                TypoDescender = (short)ReadU16(bytes, 70);
                TypoLineGap = (short)ReadU16(bytes, 72);
            }
            if (bytes.Length >= 90)
                CapHeight = (short)ReadU16(bytes, 88);
            if (WeightClass >= 600)
                IsBold = true;
        }

        private void ReadNamesFromStream(Stream stream)
        {
            TableRecord name;
            if (!_tables.TryGetValue("name", out name))
                return;
            var bytes = ReadBytes(stream, name.Offset, (int)Math.Min(name.Length, 64 * 1024u));
            if (bytes == null)
                return;
            ParseNames(bytes);
        }

        private void ParseNames(byte[] bytes)
        {
            if (bytes.Length < 6)
                return;
            var count = ReadU16(bytes, 2);
            var stringOffset = ReadU16(bytes, 4);
            for (var i = 0; i < count; i++)
            {
                var rec = 6 + i * 12;
                if (rec + 12 > bytes.Length)
                    break;
                var platformId = ReadU16(bytes, rec);
                var encodingId = ReadU16(bytes, rec + 2);
                var nameId = ReadU16(bytes, rec + 6);
                var length = ReadU16(bytes, rec + 8);
                var offset = ReadU16(bytes, rec + 10);
                var start = stringOffset + offset;
                if (start + length > bytes.Length)
                    continue;

                string value;
                if (platformId == 3 || (platformId == 0) || (platformId == 2 && encodingId == 1))
                    value = Encoding.BigEndianUnicode.GetString(bytes, start, length);
                else
                    value = Encoding.ASCII.GetString(bytes, start, length);
                value = value.Trim('\0', ' ');
                if (value.Length == 0)
                    continue;

                switch (nameId)
                {
                    case 1: if (FamilyName == null || platformId == 3) FamilyName = value; break;
                    case 2: if (SubfamilyName == null || platformId == 3) SubfamilyName = value; break;
                    case 6: if (PostScriptName == null || platformId == 3) PostScriptName = value; break;
                    case 16: if (TypographicFamily == null || platformId == 3) TypographicFamily = value; break;
                }
            }

            if (SubfamilyName != null)
            {
                var sub = SubfamilyName.ToLowerInvariant();
                if (sub.Contains("bold")) IsBold = true;
                if (sub.Contains("italic") || sub.Contains("oblique")) IsItalic = true;
                if (sub == "regular" || sub == "book") { }
            }
        }

        private void Initialize(byte[] bytes, int collectionIndex)
        {
            var tag = Encoding.ASCII.GetString(bytes, 0, 4);
            if (tag == "ttcf")
            {
                var count = (int)ReadU32(bytes, 8);
                if (collectionIndex < 0 || collectionIndex >= count)
                    collectionIndex = 0;
                var fontOffset = ReadU32(bytes, 12 + collectionIndex * 4);
                _data = ExtractFromCollection(bytes, fontOffset);
            }
            else
            {
                _data = bytes;
            }

            _tables.Clear();
            var version = ReadU32(_data, 0);
            IsCff = version == 0x4F54544F;
            var numTables = ReadU16(_data, 4);
            for (var i = 0; i < numTables; i++)
            {
                var rec = 12 + i * 16;
                if (rec + 16 > _data.Length)
                    break;
                var name = Encoding.ASCII.GetString(_data, rec, 4);
                _tables[name] = new TableRecord
                {
                    Offset = ReadU32(_data, rec + 8),
                    Length = ReadU32(_data, rec + 12),
                };
            }

            ReadHead();
            ReadOs2();
            ReadPost();
            ReadMaxp();
            ReadHmtx();
            ReadCmap();
            var nameTable = GetTable("name");
            if (nameTable != null)
                ParseNames(nameTable);
        }

        /// <summary>Rebuilds a standalone SFNT font from one member of a TrueType collection.</summary>
        private static byte[] ExtractFromCollection(byte[] bytes, uint fontOffset)
        {
            var numTables = ReadU16(bytes, (int)fontOffset + 4);
            var records = new List<KeyValuePair<string, TableRecord>>();
            for (var i = 0; i < numTables; i++)
            {
                var rec = (int)fontOffset + 12 + i * 16;
                var name = Encoding.ASCII.GetString(bytes, rec, 4);
                records.Add(new KeyValuePair<string, TableRecord>(name, new TableRecord
                {
                    Offset = ReadU32(bytes, rec + 8),
                    Length = ReadU32(bytes, rec + 12),
                }));
            }

            var headerSize = 12 + records.Count * 16;
            var totalSize = headerSize;
            foreach (var r in records)
                totalSize += (int)((r.Value.Length + 3) & ~3u);

            var output = new byte[totalSize];
            Array.Copy(bytes, (int)fontOffset, output, 0, 12);

            var dataPos = headerSize;
            for (var i = 0; i < records.Count; i++)
            {
                var r = records[i];
                var rec = 12 + i * 16;
                Encoding.ASCII.GetBytes(r.Key, 0, 4, output, rec);
                // Preserve the original checksum field.
                Array.Copy(bytes, (int)fontOffset + 12 + i * 16 + 4, output, rec + 4, 4);
                WriteU32(output, rec + 8, (uint)dataPos);
                WriteU32(output, rec + 12, r.Value.Length);
                if (r.Value.Offset + r.Value.Length <= bytes.Length)
                    Array.Copy(bytes, (int)r.Value.Offset, output, dataPos, (int)r.Value.Length);
                dataPos += (int)((r.Value.Length + 3) & ~3u);
            }
            return output;
        }

        /// <summary>
        /// Builds an embeddable subset of this font: unused glyph outlines are dropped
        /// (glyph ids stay stable — the PDF maps CIDs to glyph ids via Identity) and
        /// non-rendering tables (cmap, name, GSUB/GPOS, ...) are removed. Word's own
        /// exports subset the same way; a full Arial embed is ~1MB, the subset a few KB.
        /// </summary>
        public byte[] BuildSubset(ICollection<int> usedGlyphIds)
        {
            TableRecord glyfRec, locaRec, headRec;
            if (IsCff || !_tables.TryGetValue("glyf", out glyfRec)
                || !_tables.TryGetValue("loca", out locaRec)
                || !_tables.TryGetValue("head", out headRec))
                return _data;

            var longLoca = ReadU16(_data, (int)headRec.Offset + 50) != 0;
            var numGlyphs = NumGlyphs;

            Func<int, uint> locaOf = i => longLoca
                ? ReadU32(_data, (int)locaRec.Offset + i * 4)
                : ReadU16(_data, (int)locaRec.Offset + i * 2) * 2u;

            // Composite glyphs pull in their component outlines.
            var used = new HashSet<int>(usedGlyphIds);
            used.Add(0);
            var queue = new Queue<int>(used);
            while (queue.Count > 0)
            {
                var gid = queue.Dequeue();
                if (gid < 0 || gid >= numGlyphs)
                    continue;
                var start = (int)(glyfRec.Offset + locaOf(gid));
                var end = (int)(glyfRec.Offset + locaOf(gid + 1));
                if (end <= start || end > _data.Length || start + 10 > _data.Length)
                    continue;
                var contours = (short)ReadU16(_data, start);
                if (contours >= 0)
                    continue;
                var pos = start + 10;
                while (pos + 4 <= end)
                {
                    var flags = ReadU16(_data, pos);
                    var component = ReadU16(_data, pos + 2);
                    if (used.Add(component))
                        queue.Enqueue(component);
                    pos += 4;
                    pos += (flags & 0x0001) != 0 ? 4 : 2;      // ARG_1_AND_2_ARE_WORDS
                    if ((flags & 0x0008) != 0) pos += 2;        // WE_HAVE_A_SCALE
                    else if ((flags & 0x0040) != 0) pos += 4;   // X_AND_Y_SCALE
                    else if ((flags & 0x0080) != 0) pos += 8;   // TWO_BY_TWO
                    if ((flags & 0x0020) == 0)                  // MORE_COMPONENTS
                        break;
                }
            }

            // Sparse glyf keeping ids stable: unused glyphs become zero-length entries.
            var glyf = new System.IO.MemoryStream();
            var loca = new byte[(numGlyphs + 1) * 4];
            for (var gid = 0; gid < numGlyphs; gid++)
            {
                WriteU32(loca, gid * 4, (uint)glyf.Length);
                if (!used.Contains(gid))
                    continue;
                var start = (int)(glyfRec.Offset + locaOf(gid));
                var end = (int)(glyfRec.Offset + locaOf(gid + 1));
                if (end <= start || end > _data.Length)
                    continue;
                glyf.Write(_data, start, end - start);
                while (glyf.Length % 4 != 0)
                    glyf.WriteByte(0);
            }
            WriteU32(loca, numGlyphs * 4, (uint)glyf.Length);

            var head = GetTable("head");
            if (head == null || head.Length < 52)
                return _data;
            WriteU32(head, 8, 0);              // checkSumAdjustment recomputed below
            head[50] = 0; head[51] = 1;        // long loca

            var tables = new List<KeyValuePair<string, byte[]>>();
            tables.Add(new KeyValuePair<string, byte[]>("head", head));
            tables.Add(new KeyValuePair<string, byte[]>("loca", loca));
            tables.Add(new KeyValuePair<string, byte[]>("glyf", glyf.ToArray()));
            foreach (var name in new[] { "cvt ", "fpgm", "prep", "hhea", "hmtx", "maxp", "OS/2" })
            {
                var t = GetTable(name);
                if (t != null)
                    tables.Add(new KeyValuePair<string, byte[]>(name, t));
            }
            tables.Sort((a, b) => string.CompareOrdinal(a.Key, b.Key));

            var headerSize = 12 + tables.Count * 16;
            var totalSize = headerSize;
            foreach (var t in tables)
                totalSize += (t.Value.Length + 3) & ~3;

            var output = new byte[totalSize];
            WriteU32(output, 0, 0x00010000);
            var n = tables.Count;
            var pow2 = 1;
            var selector = 0;
            while (pow2 * 2 <= n) { pow2 *= 2; selector++; }
            output[4] = (byte)(n >> 8); output[5] = (byte)n;
            var searchRange = pow2 * 16;
            output[6] = (byte)(searchRange >> 8); output[7] = (byte)searchRange;
            output[8] = (byte)(selector >> 8); output[9] = (byte)selector;
            var rangeShift = n * 16 - searchRange;
            output[10] = (byte)(rangeShift >> 8); output[11] = (byte)rangeShift;

            var dataPos = headerSize;
            var headOffset = 0;
            for (var i = 0; i < tables.Count; i++)
            {
                var t = tables[i];
                var rec = 12 + i * 16;
                Encoding.ASCII.GetBytes(t.Key, 0, 4, output, rec);
                Array.Copy(t.Value, 0, output, dataPos, t.Value.Length);
                WriteU32(output, rec + 4, TableChecksum(output, dataPos, t.Value.Length));
                WriteU32(output, rec + 8, (uint)dataPos);
                WriteU32(output, rec + 12, (uint)t.Value.Length);
                if (t.Key == "head")
                    headOffset = dataPos;
                dataPos += (t.Value.Length + 3) & ~3;
            }

            var fileSum = TableChecksum(output, 0, output.Length);
            WriteU32(output, headOffset + 8, 0xB1B0AFBAu - fileSum);
            return output;
        }

        private static uint TableChecksum(byte[] data, int offset, int length)
        {
            uint sum = 0;
            var end = offset + ((length + 3) & ~3);
            for (var i = offset; i + 3 < end; i += 4)
                sum = unchecked(sum + ReadU32(data, i));
            return sum;
        }

        private byte[] GetTable(string name)
        {
            TableRecord rec;
            if (!_tables.TryGetValue(name, out rec))
                return null;
            if (rec.Offset + rec.Length > _data.Length)
            {
                if (rec.Offset >= _data.Length)
                    return null;
                rec.Length = (uint)(_data.Length - rec.Offset);
            }
            var bytes = new byte[rec.Length];
            Array.Copy(_data, (int)rec.Offset, bytes, 0, (int)rec.Length);
            return bytes;
        }

        private void ReadHead()
        {
            var head = GetTable("head");
            if (head == null || head.Length < 54)
                return;
            UnitsPerEm = ReadU16(head, 18);
            if (UnitsPerEm == 0)
                UnitsPerEm = 1000;
            XMin = (short)ReadU16(head, 36);
            YMin = (short)ReadU16(head, 38);
            XMax = (short)ReadU16(head, 40);
            YMax = (short)ReadU16(head, 42);
            var macStyle = ReadU16(head, 44);
            IsBold = (macStyle & 1) != 0;
            IsItalic = (macStyle & 2) != 0;

            var hhea = GetTable("hhea");
            if (hhea != null && hhea.Length >= 36)
            {
                Ascender = (short)ReadU16(hhea, 4);
                Descender = (short)ReadU16(hhea, 6);
                LineGap = (short)ReadU16(hhea, 8);
                _numberOfHMetrics = ReadU16(hhea, 34);
            }
        }

        private ushort _numberOfHMetrics;

        private void ReadOs2()
        {
            var os2 = GetTable("OS/2");
            if (os2 == null || os2.Length < 64)
                return;
            WeightClass = ReadU16(os2, 4);
            var familyClass = (short)ReadU16(os2, 30);
            IsSerif = (familyClass >> 8) >= 1 && (familyClass >> 8) <= 7;
            var fsSelection = ReadU16(os2, 62);
            if ((fsSelection & 0x20) != 0) IsBold = true;
            if ((fsSelection & 0x01) != 0) IsItalic = true;
            UseTypoMetrics = (fsSelection & 0x80) != 0;
            if (os2.Length >= 74)
            {
                TypoAscender = (short)ReadU16(os2, 68);
                TypoDescender = (short)ReadU16(os2, 70);
                TypoLineGap = (short)ReadU16(os2, 72);
            }
            if (os2.Length >= 78)
            {
                WinAscent = ReadU16(os2, 74);
                WinDescent = ReadU16(os2, 76);
            }
            if (os2.Length >= 90)
                CapHeight = (short)ReadU16(os2, 88);
        }

        private void ReadPost()
        {
            var post = GetTable("post");
            if (post == null || post.Length < 32)
                return;
            var raw = (int)ReadU32(post, 4);
            ItalicAngle = raw / 65536.0;
            IsFixedPitch = ReadU32(post, 12) != 0;
        }

        private void ReadMaxp()
        {
            var maxp = GetTable("maxp");
            if (maxp != null && maxp.Length >= 6)
                NumGlyphs = ReadU16(maxp, 4);
        }

        private void ReadHmtx()
        {
            var hmtx = GetTable("hmtx");
            if (hmtx == null || _numberOfHMetrics == 0)
                return;
            var count = Math.Min((int)_numberOfHMetrics, hmtx.Length / 4);
            _advanceWidths = new ushort[Math.Max(count, 1)];
            for (var i = 0; i < count; i++)
                _advanceWidths[i] = ReadU16(hmtx, i * 4);
        }

        private void ReadCmap()
        {
            _cmap = new Dictionary<int, int>();
            var cmap = GetTable("cmap");
            if (cmap == null || cmap.Length < 4)
                return;

            var numTables = ReadU16(cmap, 2);
            int best = -1, bestScore = -1;
            int symbol = -1;
            for (var i = 0; i < numTables; i++)
            {
                var rec = 4 + i * 8;
                if (rec + 8 > cmap.Length)
                    break;
                var platform = ReadU16(cmap, rec);
                var encoding = ReadU16(cmap, rec + 2);
                var offset = (int)ReadU32(cmap, rec + 4);
                if (offset <= 0 || offset >= cmap.Length)
                    continue;

                var score = -1;
                if (platform == 3 && encoding == 10) score = 100;
                else if (platform == 0 && encoding >= 4) score = 95;
                else if (platform == 3 && encoding == 1) score = 90;
                else if (platform == 0) score = 80;
                else if (platform == 3 && encoding == 0) { symbol = offset; score = 10; }
                else if (platform == 1 && encoding == 0) score = 5;

                if (score > bestScore)
                {
                    bestScore = score;
                    best = offset;
                }
            }

            if (best >= 0)
                ParseCmapSubtable(cmap, best, _cmap);
            if (symbol >= 0)
            {
                _symbolCmap = new Dictionary<int, int>();
                ParseCmapSubtable(cmap, symbol, _symbolCmap);
            }
        }

        private static void ParseCmapSubtable(byte[] cmap, int offset, Dictionary<int, int> target)
        {
            if (offset + 4 > cmap.Length)
                return;
            var format = ReadU16(cmap, offset);
            switch (format)
            {
                case 0:
                {
                    for (var i = 0; i < 256; i++)
                    {
                        var idx = offset + 6 + i;
                        if (idx >= cmap.Length)
                            break;
                        var gid = cmap[idx];
                        if (gid != 0)
                            target[i] = gid;
                    }
                    break;
                }
                case 4:
                {
                    var segCountX2 = ReadU16(cmap, offset + 6);
                    var segCount = segCountX2 / 2;
                    var endBase = offset + 14;
                    var startBase = endBase + segCountX2 + 2;
                    var deltaBase = startBase + segCountX2;
                    var rangeBase = deltaBase + segCountX2;
                    for (var seg = 0; seg < segCount; seg++)
                    {
                        if (rangeBase + seg * 2 + 2 > cmap.Length)
                            break;
                        int end = ReadU16(cmap, endBase + seg * 2);
                        int start = ReadU16(cmap, startBase + seg * 2);
                        int delta = (short)ReadU16(cmap, deltaBase + seg * 2);
                        int rangeOffset = ReadU16(cmap, rangeBase + seg * 2);
                        if (start > end)
                            continue;
                        for (var c = start; c <= end && c != 0xFFFF; c++)
                        {
                            int gid;
                            if (rangeOffset == 0)
                            {
                                gid = (c + delta) & 0xFFFF;
                            }
                            else
                            {
                                var glyphIndexAddress = rangeBase + seg * 2 + rangeOffset + (c - start) * 2;
                                if (glyphIndexAddress + 2 > cmap.Length)
                                    continue;
                                gid = ReadU16(cmap, glyphIndexAddress);
                                if (gid != 0)
                                    gid = (gid + delta) & 0xFFFF;
                            }
                            if (gid != 0 && !target.ContainsKey(c))
                                target[c] = gid;
                        }
                    }
                    break;
                }
                case 6:
                {
                    var first = ReadU16(cmap, offset + 6);
                    var count = ReadU16(cmap, offset + 8);
                    for (var i = 0; i < count; i++)
                    {
                        var idx = offset + 10 + i * 2;
                        if (idx + 2 > cmap.Length)
                            break;
                        var gid = ReadU16(cmap, idx);
                        if (gid != 0)
                            target[first + i] = gid;
                    }
                    break;
                }
                case 12:
                {
                    var nGroups = (int)ReadU32(cmap, offset + 12);
                    for (var g = 0; g < nGroups; g++)
                    {
                        var rec = offset + 16 + g * 12;
                        if (rec + 12 > cmap.Length)
                            break;
                        var startChar = (int)ReadU32(cmap, rec);
                        var endChar = (int)ReadU32(cmap, rec + 4);
                        var startGid = (int)ReadU32(cmap, rec + 8);
                        if (endChar - startChar > 0x20000)
                            endChar = startChar + 0x20000;
                        for (var c = startChar; c <= endChar; c++)
                        {
                            if (!target.ContainsKey(c))
                                target[c] = startGid + (c - startChar);
                        }
                    }
                    break;
                }
            }
        }

        /// <summary>Glyph id for a code point, or 0 when the font has no such glyph.</summary>
        public int GetGlyphId(int codePoint)
        {
            int gid;
            if (_cmap != null && _cmap.TryGetValue(codePoint, out gid))
                return gid;
            if (_symbolCmap != null)
            {
                if (_symbolCmap.TryGetValue(codePoint, out gid))
                    return gid;
                // Symbol fonts map their glyphs into the F0xx private use range.
                if (codePoint < 0x100 && _symbolCmap.TryGetValue(0xF000 + codePoint, out gid))
                    return gid;
                if (codePoint >= 0xF000 && codePoint <= 0xF0FF && _symbolCmap.TryGetValue(codePoint & 0xFF, out gid))
                    return gid;
            }
            if (_cmap != null && codePoint >= 0xF000 && codePoint <= 0xF0FF
                && _cmap.TryGetValue(codePoint & 0xFF, out gid))
                return gid;
            return 0;
        }

        /// <summary>Advance width of a glyph in font design units.</summary>
        public int GetAdvance(int glyphId)
        {
            if (_advanceWidths == null || _advanceWidths.Length == 0)
                return UnitsPerEm / 2;
            if (glyphId < 0)
                return 0;
            return glyphId < _advanceWidths.Length
                ? _advanceWidths[glyphId]
                : _advanceWidths[_advanceWidths.Length - 1];
        }

        public bool HasGlyphOutlines
        {
            get { return _tables.ContainsKey("glyf") || _tables.ContainsKey("CFF "); }
        }

        private static ushort ReadU16(byte[] b, int offset)
        {
            return (ushort)((b[offset] << 8) | b[offset + 1]);
        }

        private static uint ReadU32(byte[] b, int offset)
        {
            return ((uint)b[offset] << 24) | ((uint)b[offset + 1] << 16) | ((uint)b[offset + 2] << 8) | b[offset + 3];
        }

        private static void WriteU32(byte[] b, int offset, uint value)
        {
            b[offset] = (byte)(value >> 24);
            b[offset + 1] = (byte)(value >> 16);
            b[offset + 2] = (byte)(value >> 8);
            b[offset + 3] = (byte)value;
        }
    }
}
