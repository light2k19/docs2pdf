using System;
using System.IO;
using System.IO.Compression;

namespace Docx2Pdf.Pdf
{
    /// <summary>zlib (RFC 1950) wrapper around the BCL raw-deflate implementation, as required by /FlateDecode.</summary>
    internal static class Flate
    {
        public static byte[] Compress(byte[] data)
        {
            if (data == null)
                return new byte[0];

            using (var output = new MemoryStream())
            {
                output.WriteByte(0x78);          // CMF: deflate, 32K window
                output.WriteByte(0x9C);          // FLG: default compression, no dictionary
                using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true))
                    deflate.Write(data, 0, data.Length);

                var adler = Adler32(data);
                output.WriteByte((byte)(adler >> 24));
                output.WriteByte((byte)(adler >> 16));
                output.WriteByte((byte)(adler >> 8));
                output.WriteByte((byte)adler);
                return output.ToArray();
            }
        }

        /// <summary>Inflates a zlib stream (used for PNG IDAT data).</summary>
        public static byte[] Decompress(byte[] data, int offset, int count)
        {
            if (data == null || count <= 0)
                return new byte[0];

            // Skip the 2-byte zlib header when present.
            var start = offset;
            var length = count;
            if (count >= 2)
            {
                var cmf = data[offset];
                var flg = data[offset + 1];
                if ((cmf & 0x0F) == 8 && ((cmf << 8) | flg) % 31 == 0)
                {
                    start += 2;
                    length -= 2;
                    if ((flg & 0x20) != 0)     // FDICT
                    {
                        start += 4;
                        length -= 4;
                    }
                }
            }
            if (length <= 0)
                return new byte[0];

            using (var input = new MemoryStream(data, start, length, false))
            using (var deflate = new DeflateStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                var buffer = new byte[16384];
                int read;
                while ((read = deflate.Read(buffer, 0, buffer.Length)) > 0)
                    output.Write(buffer, 0, read);
                return output.ToArray();
            }
        }

        public static uint Adler32(byte[] data)
        {
            const uint mod = 65521;
            uint a = 1, b = 0;
            foreach (var t in data)
            {
                a = (a + t) % mod;
                b = (b + a) % mod;
            }
            return (b << 16) | a;
        }
    }
}
