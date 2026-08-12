using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Docx2Pdf.Pdf
{
    internal abstract class PdfObject
    {
        public abstract void Write(Stream stream);

        protected static void Ascii(Stream stream, string text)
        {
            var bytes = Encoding.ASCII.GetBytes(text);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    internal sealed class PdfNull : PdfObject
    {
        public static readonly PdfNull Instance = new PdfNull();
        public override void Write(Stream stream) { Ascii(stream, "null"); }
    }

    internal sealed class PdfBool : PdfObject
    {
        public static readonly PdfBool True = new PdfBool(true);
        public static readonly PdfBool False = new PdfBool(false);
        private readonly bool _value;
        private PdfBool(bool value) { _value = value; }
        public static PdfBool Of(bool value) { return value ? True : False; }
        public override void Write(Stream stream) { Ascii(stream, _value ? "true" : "false"); }
    }

    internal sealed class PdfNumber : PdfObject
    {
        private readonly double _value;
        public PdfNumber(double value) { _value = value; }
        public override void Write(Stream stream) { Ascii(stream, Format(_value)); }

        public static string Format(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                value = 0;
            if (Math.Abs(value - Math.Round(value)) < 1e-6)
                return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
            var s = value.ToString("0.####", CultureInfo.InvariantCulture);
            return s.Length == 0 || s == "-0" ? "0" : s;
        }
    }

    internal sealed class PdfName : PdfObject
    {
        public readonly string Value;
        public PdfName(string value) { Value = value; }
        public override void Write(Stream stream)
        {
            var sb = new StringBuilder("/");
            foreach (var c in Value)
            {
                if (c < 33 || c > 126 || c == '#' || c == '/' || c == '(' || c == ')' || c == '<' || c == '>'
                    || c == '[' || c == ']' || c == '{' || c == '}' || c == '%')
                    sb.Append('#').Append(((int)c).ToString("X2", CultureInfo.InvariantCulture));
                else
                    sb.Append(c);
            }
            Ascii(stream, sb.ToString());
        }
    }

    /// <summary>A PDF text string. Encoded as UTF-16BE hex when it contains non-ASCII characters.</summary>
    internal sealed class PdfString : PdfObject
    {
        private readonly string _value;
        private readonly bool _rawHex;

        public PdfString(string value) { _value = value ?? string.Empty; }
        private PdfString(string hex, bool raw) { _value = hex; _rawHex = raw; }

        public static PdfString FromHex(string hexDigits) { return new PdfString(hexDigits, true); }

        public override void Write(Stream stream)
        {
            if (_rawHex)
            {
                Ascii(stream, "<" + _value + ">");
                return;
            }

            var ascii = true;
            foreach (var c in _value)
            {
                if (c > 126)
                {
                    ascii = false;
                    break;
                }
            }

            if (ascii)
            {
                var sb = new StringBuilder("(");
                foreach (var c in _value)
                {
                    if (c == '(' || c == ')' || c == '\\')
                        sb.Append('\\').Append(c);
                    else if (c == '\r')
                        sb.Append("\\r");
                    else if (c == '\n')
                        sb.Append("\\n");
                    else if (c == '\t')
                        sb.Append("\\t");
                    else if (c < 32)
                        sb.Append('\\').Append(Convert.ToString(c, 8).PadLeft(3, '0'));
                    else
                        sb.Append(c);
                }
                sb.Append(')');
                Ascii(stream, sb.ToString());
            }
            else
            {
                var sb = new StringBuilder("<FEFF");
                foreach (var c in _value)
                    sb.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                sb.Append('>');
                Ascii(stream, sb.ToString());
            }
        }
    }

    internal sealed class PdfArray : PdfObject
    {
        public readonly List<PdfObject> Items = new List<PdfObject>();

        public PdfArray() { }
        public PdfArray(params PdfObject[] items) { Items.AddRange(items); }

        public PdfArray Add(PdfObject item) { Items.Add(item); return this; }
        public PdfArray Add(double number) { Items.Add(new PdfNumber(number)); return this; }

        public override void Write(Stream stream)
        {
            Ascii(stream, "[");
            for (var i = 0; i < Items.Count; i++)
            {
                if (i > 0)
                    Ascii(stream, " ");
                Items[i].Write(stream);
            }
            Ascii(stream, "]");
        }
    }

    internal class PdfDictionary : PdfObject
    {
        public readonly List<KeyValuePair<string, PdfObject>> Entries = new List<KeyValuePair<string, PdfObject>>();

        public PdfDictionary Set(string key, PdfObject value)
        {
            for (var i = 0; i < Entries.Count; i++)
            {
                if (Entries[i].Key == key)
                {
                    Entries[i] = new KeyValuePair<string, PdfObject>(key, value);
                    return this;
                }
            }
            Entries.Add(new KeyValuePair<string, PdfObject>(key, value));
            return this;
        }

        public PdfDictionary Set(string key, string name) { return Set(key, new PdfName(name)); }
        public PdfDictionary Set(string key, double number) { return Set(key, new PdfNumber(number)); }

        public bool Contains(string key)
        {
            foreach (var e in Entries)
            {
                if (e.Key == key)
                    return true;
            }
            return false;
        }

        public override void Write(Stream stream)
        {
            Ascii(stream, "<<");
            foreach (var entry in Entries)
            {
                new PdfName(entry.Key).Write(stream);
                Ascii(stream, " ");
                entry.Value.Write(stream);
                Ascii(stream, " ");
            }
            Ascii(stream, ">>");
        }
    }

    internal sealed class PdfStream : PdfDictionary
    {
        public byte[] Data;

        public PdfStream(byte[] data) { Data = data ?? new byte[0]; }

        public override void Write(Stream stream)
        {
            Set("Length", new PdfNumber(Data.Length));
            base.Write(stream);
            Ascii(stream, "\nstream\n");
            stream.Write(Data, 0, Data.Length);
            Ascii(stream, "\nendstream");
        }
    }

    internal sealed class PdfReference : PdfObject
    {
        public readonly int Number;
        public readonly int Generation;
        public PdfObject Target;

        public PdfReference(int number, int generation, PdfObject target)
        {
            Number = number;
            Generation = generation;
            Target = target;
        }

        public override void Write(Stream stream)
        {
            Ascii(stream, Number.ToString(CultureInfo.InvariantCulture) + " "
                        + Generation.ToString(CultureInfo.InvariantCulture) + " R");
        }
    }
}
