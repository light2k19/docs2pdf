using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Docx2Pdf.Pdf
{
    /// <summary>Collects indirect objects and serialises a complete PDF file with a classic xref table.</summary>
    internal sealed class PdfDocument
    {
        private readonly List<PdfReference> _objects = new List<PdfReference>();

        public PdfDictionary Catalog;
        public PdfDictionary Info;
        public string FileId = "Docx2Pdf";

        public PdfReference Add(PdfObject obj)
        {
            var reference = new PdfReference(_objects.Count + 1, 0, obj);
            _objects.Add(reference);
            return reference;
        }

        /// <summary>Reserves an object number so it can be referenced before its body is known.</summary>
        public PdfReference Reserve()
        {
            return Add(PdfNull.Instance);
        }

        public void Save(Stream stream)
        {
            var offsets = new long[_objects.Count + 1];
            var writer = new CountingStream(stream);

            WriteAscii(writer, "%PDF-1.7\n");
            writer.Write(new byte[] { (byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n' }, 0, 6);

            foreach (var reference in _objects)
            {
                offsets[reference.Number] = writer.Position;
                WriteAscii(writer, reference.Number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
                reference.Target.Write(writer);
                WriteAscii(writer, "\nendobj\n");
            }

            var xrefPos = writer.Position;
            WriteAscii(writer, "xref\n0 " + (_objects.Count + 1).ToString(CultureInfo.InvariantCulture) + "\n");
            WriteAscii(writer, "0000000000 65535 f \n");
            for (var i = 1; i <= _objects.Count; i++)
                WriteAscii(writer, offsets[i].ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");

            var trailer = new PdfDictionary();
            trailer.Set("Size", new PdfNumber(_objects.Count + 1));
            if (Catalog != null)
                trailer.Set("Root", FindReference(Catalog));
            if (Info != null)
                trailer.Set("Info", FindReference(Info));
            var id = Hash(FileId + "|" + _objects.Count);
            trailer.Set("ID", new PdfArray(PdfString.FromHex(id), PdfString.FromHex(id)));

            WriteAscii(writer, "trailer\n");
            trailer.Write(writer);
            WriteAscii(writer, "\nstartxref\n" + xrefPos.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
            writer.Flush();
        }

        private PdfObject FindReference(PdfObject target)
        {
            foreach (var reference in _objects)
            {
                if (ReferenceEquals(reference.Target, target))
                    return reference;
            }
            return PdfNull.Instance;
        }

        private static string Hash(string input)
        {
            // Deterministic 16-byte identifier (FNV-1a based); no crypto dependency needed.
            unchecked
            {
                var sb = new StringBuilder();
                ulong h1 = 14695981039346656037;
                foreach (var c in input)
                {
                    h1 ^= c;
                    h1 *= 1099511628211;
                }
                var h2 = h1 * 3141592653589793239UL + 2718281828459045235UL;
                sb.Append(h1.ToString("X16", CultureInfo.InvariantCulture));
                sb.Append(h2.ToString("X16", CultureInfo.InvariantCulture));
                return sb.ToString();
            }
        }

        private static void WriteAscii(Stream stream, string text)
        {
            var bytes = Encoding.ASCII.GetBytes(text);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>Wraps an output stream so byte offsets are known even for non-seekable streams.</summary>
        private sealed class CountingStream : Stream
        {
            private readonly Stream _inner;
            private long _position;

            public CountingStream(Stream inner) { _inner = inner; }

            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { return _position; } }
            public override long Position
            {
                get { return _position; }
                set { throw new NotSupportedException(); }
            }

            public override void Flush() { _inner.Flush(); }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }

            public override void Write(byte[] buffer, int offset, int count)
            {
                _inner.Write(buffer, offset, count);
                _position += count;
            }

            public override void WriteByte(byte value)
            {
                _inner.WriteByte(value);
                _position++;
            }
        }
    }
}
