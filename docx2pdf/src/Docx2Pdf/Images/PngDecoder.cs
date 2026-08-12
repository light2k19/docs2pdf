using System;
using System.Collections.Generic;
using System.IO;
using Docx2Pdf.Pdf;

namespace Docx2Pdf.Images
{
    /// <summary>Full PNG decoder: all colour types and bit depths, including Adam7 interlacing.</summary>
    internal static class PngDecoder
    {
        private static readonly byte[] Signature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

        public static bool IsPng(byte[] data)
        {
            if (data == null || data.Length < 8)
                return false;
            for (var i = 0; i < 8; i++)
            {
                if (data[i] != Signature[i])
                    return false;
            }
            return true;
        }

        public static DecodedImage Decode(byte[] data)
        {
            if (!IsPng(data))
                return null;

            int width = 0, height = 0, bitDepth = 8, colorType = 6, interlace = 0;
            byte[] palette = null;
            byte[] paletteAlpha = null;
            int[] transparentColor = null;
            var idat = new MemoryStream();

            var pos = 8;
            while (pos + 8 <= data.Length)
            {
                var length = (int)ReadU32(data, pos);
                var type = System.Text.Encoding.ASCII.GetString(data, pos + 4, 4);
                var dataStart = pos + 8;
                if (length < 0 || dataStart + length > data.Length)
                    break;

                switch (type)
                {
                    case "IHDR":
                        if (length < 13)
                            return null;
                        width = (int)ReadU32(data, dataStart);
                        height = (int)ReadU32(data, dataStart + 4);
                        bitDepth = data[dataStart + 8];
                        colorType = data[dataStart + 9];
                        interlace = data[dataStart + 12];
                        break;
                    case "PLTE":
                        palette = new byte[length];
                        Array.Copy(data, dataStart, palette, 0, length);
                        break;
                    case "tRNS":
                        if (colorType == 3)
                        {
                            paletteAlpha = new byte[length];
                            Array.Copy(data, dataStart, paletteAlpha, 0, length);
                        }
                        else if (colorType == 0 && length >= 2)
                        {
                            transparentColor = new[] { ReadU16(data, dataStart) };
                        }
                        else if (colorType == 2 && length >= 6)
                        {
                            transparentColor = new[]
                            {
                                ReadU16(data, dataStart),
                                ReadU16(data, dataStart + 2),
                                ReadU16(data, dataStart + 4),
                            };
                        }
                        break;
                    case "IDAT":
                        idat.Write(data, dataStart, length);
                        break;
                    case "IEND":
                        pos = data.Length;
                        break;
                }

                pos = dataStart + length + 4;
            }

            if (width <= 0 || height <= 0 || (long)width * height > 80_000_000L)
                return null;

            var compressed = idat.ToArray();
            if (compressed.Length == 0)
                return null;

            byte[] raw;
            try
            {
                raw = Flate.Decompress(compressed, 0, compressed.Length);
            }
            catch
            {
                return null;
            }
            if (raw.Length == 0)
                return null;

            var channels = ChannelsOf(colorType);
            if (channels == 0)
                return null;

            // Decode to 8-bit-per-channel samples in the source colour model.
            var sampleCount = channels;
            var pixels = new byte[(long)width * height * sampleCount > int.MaxValue ? 0 : width * height * sampleCount];
            if (pixels.Length == 0)
                return null;

            if (interlace == 1)
            {
                if (!DecodeAdam7(raw, width, height, bitDepth, channels, pixels))
                    return null;
            }
            else
            {
                if (!DecodePass(raw, 0, width, height, bitDepth, channels, pixels, width, 0, 0, 1, 1))
                    return null;
            }

            return BuildImage(width, height, colorType, bitDepth, channels, pixels, palette, paletteAlpha, transparentColor);
        }

        private static int ChannelsOf(int colorType)
        {
            switch (colorType)
            {
                case 0: return 1;
                case 2: return 3;
                case 3: return 1;
                case 4: return 2;
                case 6: return 4;
                default: return 0;
            }
        }

        private static bool DecodeAdam7(byte[] raw, int width, int height, int bitDepth, int channels, byte[] output)
        {
            int[] xStart = { 0, 4, 0, 2, 0, 1, 0 };
            int[] yStart = { 0, 0, 4, 0, 2, 0, 1 };
            int[] xStep = { 8, 8, 4, 4, 2, 2, 1 };
            int[] yStep = { 8, 8, 8, 4, 4, 2, 2 };

            var offset = 0;
            for (var pass = 0; pass < 7; pass++)
            {
                var passWidth = (width - xStart[pass] + xStep[pass] - 1) / xStep[pass];
                var passHeight = (height - yStart[pass] + yStep[pass] - 1) / yStep[pass];
                if (passWidth <= 0 || passHeight <= 0)
                    continue;

                if (!DecodePass(raw, offset, passWidth, passHeight, bitDepth, channels, output, width,
                                xStart[pass], yStart[pass], xStep[pass], yStep[pass]))
                    return false;

                var rowBytes = (passWidth * channels * bitDepth + 7) / 8;
                offset += (rowBytes + 1) * passHeight;
            }
            return true;
        }

        /// <summary>Unfilters one (sub)image and scatters its pixels into the output buffer.</summary>
        private static bool DecodePass(byte[] raw, int offset, int passWidth, int passHeight, int bitDepth,
                                       int channels, byte[] output, int imageWidth,
                                       int xStart, int yStart, int xStep, int yStep)
        {
            var rowBytes = (passWidth * channels * bitDepth + 7) / 8;
            var bpp = Math.Max(1, channels * bitDepth / 8);
            var previous = new byte[rowBytes];
            var current = new byte[rowBytes];

            for (var y = 0; y < passHeight; y++)
            {
                var rowStart = offset + y * (rowBytes + 1);
                if (rowStart + rowBytes + 1 > raw.Length)
                    return y > 0;         // truncated data: keep what was decoded

                var filter = raw[rowStart];
                Array.Copy(raw, rowStart + 1, current, 0, rowBytes);
                Unfilter(filter, current, previous, bpp);

                // Expand to 8 bits per sample and place into the target rows.
                var targetY = yStart + y * yStep;
                if (targetY >= (output.Length / channels) / imageWidth)
                    continue;

                for (var x = 0; x < passWidth; x++)
                {
                    var targetX = xStart + x * xStep;
                    if (targetX >= imageWidth)
                        break;
                    var target = (targetY * imageWidth + targetX) * channels;
                    if (target + channels > output.Length)
                        break;

                    for (var c = 0; c < channels; c++)
                    {
                        var sampleIndex = x * channels + c;
                        output[target + c] = ReadSample(current, sampleIndex, bitDepth);
                    }
                }

                var swap = previous;
                previous = current;
                current = swap;
            }
            return true;
        }

        private static byte ReadSample(byte[] row, int index, int bitDepth)
        {
            switch (bitDepth)
            {
                case 8:
                    return index < row.Length ? row[index] : (byte)0;
                case 16:
                {
                    var i = index * 2;
                    return i < row.Length ? row[i] : (byte)0;
                }
                case 1:
                case 2:
                case 4:
                {
                    var bitsPerSample = bitDepth;
                    var bitPos = index * bitsPerSample;
                    var byteIndex = bitPos / 8;
                    if (byteIndex >= row.Length)
                        return 0;
                    var shift = 8 - bitsPerSample - (bitPos % 8);
                    var mask = (1 << bitsPerSample) - 1;
                    var value = (row[byteIndex] >> shift) & mask;
                    return (byte)value;
                }
                default:
                    return 0;
            }
        }

        private static void Unfilter(int filter, byte[] current, byte[] previous, int bpp)
        {
            switch (filter)
            {
                case 0:
                    break;
                case 1:
                    for (var i = bpp; i < current.Length; i++)
                        current[i] = (byte)(current[i] + current[i - bpp]);
                    break;
                case 2:
                    for (var i = 0; i < current.Length; i++)
                        current[i] = (byte)(current[i] + previous[i]);
                    break;
                case 3:
                    for (var i = 0; i < current.Length; i++)
                    {
                        var left = i >= bpp ? current[i - bpp] : 0;
                        current[i] = (byte)(current[i] + ((left + previous[i]) >> 1));
                    }
                    break;
                case 4:
                    for (var i = 0; i < current.Length; i++)
                    {
                        var a = i >= bpp ? current[i - bpp] : 0;
                        var b = previous[i];
                        var c = i >= bpp ? previous[i - bpp] : 0;
                        current[i] = (byte)(current[i] + Paeth(a, b, c));
                    }
                    break;
            }
        }

        private static int Paeth(int a, int b, int c)
        {
            var p = a + b - c;
            var pa = Math.Abs(p - a);
            var pb = Math.Abs(p - b);
            var pc = Math.Abs(p - c);
            if (pa <= pb && pa <= pc)
                return a;
            return pb <= pc ? b : c;
        }

        private static DecodedImage BuildImage(int width, int height, int colorType, int bitDepth, int channels,
                                               byte[] pixels, byte[] palette, byte[] paletteAlpha, int[] transparent)
        {
            var pixelCount = width * height;
            var maxValue = (1 << Math.Min(bitDepth, 8)) - 1;

            switch (colorType)
            {
                case 0:   // grayscale
                {
                    var gray = new byte[pixelCount];
                    byte[] alpha = null;
                    for (var i = 0; i < pixelCount; i++)
                        gray[i] = bitDepth < 8 ? (byte)(pixels[i] * 255 / maxValue) : pixels[i];

                    if (transparent != null)
                    {
                        var key = bitDepth <= 8 ? transparent[0] : transparent[0] >> 8;
                        var keyByte = bitDepth < 8 ? (byte)(key * 255 / maxValue) : (byte)key;
                        alpha = new byte[pixelCount];
                        for (var i = 0; i < pixelCount; i++)
                            alpha[i] = gray[i] == keyByte ? (byte)0 : (byte)255;
                    }
                    return new DecodedImage { Width = width, Height = height, Components = 1, Samples = gray, Alpha = alpha };
                }
                case 4:   // grayscale + alpha
                {
                    var gray = new byte[pixelCount];
                    var alpha = new byte[pixelCount];
                    for (var i = 0; i < pixelCount; i++)
                    {
                        gray[i] = pixels[i * 2];
                        alpha[i] = pixels[i * 2 + 1];
                    }
                    return new DecodedImage { Width = width, Height = height, Components = 1, Samples = gray, Alpha = alpha };
                }
                case 2:   // truecolour
                {
                    var rgb = new byte[pixelCount * 3];
                    Array.Copy(pixels, rgb, Math.Min(pixels.Length, rgb.Length));
                    byte[] alpha = null;
                    if (transparent != null && transparent.Length >= 3)
                    {
                        var kr = (byte)(bitDepth <= 8 ? transparent[0] : transparent[0] >> 8);
                        var kg = (byte)(bitDepth <= 8 ? transparent[1] : transparent[1] >> 8);
                        var kb = (byte)(bitDepth <= 8 ? transparent[2] : transparent[2] >> 8);
                        alpha = new byte[pixelCount];
                        for (var i = 0; i < pixelCount; i++)
                        {
                            var opaque = !(rgb[i * 3] == kr && rgb[i * 3 + 1] == kg && rgb[i * 3 + 2] == kb);
                            alpha[i] = opaque ? (byte)255 : (byte)0;
                        }
                    }
                    return DecodedImage.FromRgb(width, height, rgb, alpha);
                }
                case 6:   // truecolour + alpha
                {
                    var rgb = new byte[pixelCount * 3];
                    var alpha = new byte[pixelCount];
                    for (var i = 0; i < pixelCount; i++)
                    {
                        rgb[i * 3] = pixels[i * 4];
                        rgb[i * 3 + 1] = pixels[i * 4 + 1];
                        rgb[i * 3 + 2] = pixels[i * 4 + 2];
                        alpha[i] = pixels[i * 4 + 3];
                    }
                    return DecodedImage.FromRgb(width, height, rgb, alpha);
                }
                case 3:   // indexed
                {
                    if (palette == null)
                        return null;
                    var rgb = new byte[pixelCount * 3];
                    byte[] alpha = paletteAlpha != null ? new byte[pixelCount] : null;
                    var entries = palette.Length / 3;
                    for (var i = 0; i < pixelCount; i++)
                    {
                        var index = pixels[i];
                        if (index >= entries)
                            index = 0;
                        rgb[i * 3] = palette[index * 3];
                        rgb[i * 3 + 1] = palette[index * 3 + 1];
                        rgb[i * 3 + 2] = palette[index * 3 + 2];
                        if (alpha != null)
                            alpha[i] = index < paletteAlpha.Length ? paletteAlpha[index] : (byte)255;
                    }
                    return DecodedImage.FromRgb(width, height, rgb, alpha);
                }
                default:
                    return null;
            }
        }

        private static uint ReadU32(byte[] b, int offset)
        {
            return ((uint)b[offset] << 24) | ((uint)b[offset + 1] << 16) | ((uint)b[offset + 2] << 8) | b[offset + 3];
        }

        private static int ReadU16(byte[] b, int offset)
        {
            return (b[offset] << 8) | b[offset + 1];
        }
    }
}
