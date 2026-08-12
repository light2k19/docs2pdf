using System;
using System.Collections.Generic;

namespace Docx2Pdf.Images
{
    /// <summary>Recognises the image formats found in .docx packages and decodes them for PDF embedding.</summary>
    internal static class ImageDecoder
    {
        /// <summary>Decodes an image, or returns null when the format is not supported.</summary>
        public static DecodedImage Decode(byte[] data, out string formatName)
        {
            formatName = "unknown";
            if (data == null || data.Length < 8)
                return null;

            if (PngDecoder.IsPng(data))
            {
                formatName = "PNG";
                return PngDecoder.Decode(data);
            }
            if (data[0] == 0xFF && data[1] == 0xD8)
            {
                formatName = "JPEG";
                return DecodeJpeg(data);
            }
            if (data[0] == 'B' && data[1] == 'M')
            {
                formatName = "BMP";
                return DecodeBmp(data);
            }
            if (data[0] == 'G' && data[1] == 'I' && data[2] == 'F')
            {
                formatName = "GIF";
                return GifDecoder.Decode(data);
            }
            if ((data[0] == 'I' && data[1] == 'I' && data[2] == 0x2A) || (data[0] == 'M' && data[1] == 'M' && data[3] == 0x2A))
                formatName = "TIFF";
            else if (data[0] == 0xD7 && data[1] == 0xCD)
                formatName = "WMF";
            else if (data.Length > 40 && data[40] == 0x20 && data[41] == 0x45 && data[42] == 0x4D && data[43] == 0x46)
                formatName = "EMF";
            else if (data[0] == 'I' && data[1] == 'I' && data[2] == 0xBC)
                formatName = "WDP";
            return null;
        }

        // ------------------------------------------------------------------ JPEG

        /// <summary>JPEG data is embedded as-is; only the frame header is parsed.</summary>
        private static DecodedImage DecodeJpeg(byte[] data)
        {
            var pos = 2;
            var adobeInverted = false;
            while (pos + 4 <= data.Length)
            {
                if (data[pos] != 0xFF)
                {
                    pos++;
                    continue;
                }

                var marker = data[pos + 1];
                if (marker == 0xD8 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
                {
                    pos += 2;
                    continue;
                }
                if (marker == 0xD9 || marker == 0xDA)
                    break;

                if (pos + 4 > data.Length)
                    break;
                var length = (data[pos + 2] << 8) | data[pos + 3];
                if (length < 2 || pos + 2 + length > data.Length)
                    break;

                var isSof = marker >= 0xC0 && marker <= 0xCF
                            && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                if (isSof && length >= 8)
                {
                    var height = (data[pos + 5] << 8) | data[pos + 6];
                    var width = (data[pos + 7] << 8) | data[pos + 8];
                    var components = data[pos + 9];
                    if (width > 0 && height > 0)
                    {
                        return new DecodedImage
                        {
                            Width = width,
                            Height = height,
                            Components = components == 1 ? 1 : 3,
                            JpegData = data,
                            JpegComponents = components,
                            JpegAdobeInverted = adobeInverted && components == 4,
                        };
                    }
                }

                if (marker == 0xEE && length >= 12)
                {
                    // Adobe APP14: CMYK data written by Photoshop is stored inverted.
                    var isAdobe = data[pos + 4] == 'A' && data[pos + 5] == 'd' && data[pos + 6] == 'o';
                    if (isAdobe)
                        adobeInverted = true;
                }

                pos += 2 + length;
            }
            return null;
        }

        // ------------------------------------------------------------------ BMP

        private static DecodedImage DecodeBmp(byte[] data)
        {
            if (data.Length < 26)
                return null;

            var pixelOffset = (int)ReadU32Le(data, 10);
            var headerSize = (int)ReadU32Le(data, 14);

            int width, height, bitCount, compression = 0, paletteEntries = 0;
            var topDown = false;
            var paletteEntrySize = 4;

            if (headerSize == 12)
            {
                width = ReadU16Le(data, 18);
                height = ReadU16Le(data, 20);
                bitCount = ReadU16Le(data, 24);
                paletteEntrySize = 3;
            }
            else if (headerSize >= 40 && data.Length >= 54)
            {
                width = (int)ReadU32Le(data, 18);
                height = (int)ReadU32Le(data, 22);
                bitCount = ReadU16Le(data, 28);
                compression = (int)ReadU32Le(data, 30);
                paletteEntries = (int)ReadU32Le(data, 46);
                if (height < 0)
                {
                    height = -height;
                    topDown = true;
                }
            }
            else
            {
                return null;
            }

            if (width <= 0 || height <= 0 || (long)width * height > 80_000_000L)
                return null;
            if (compression != 0 && compression != 3)
                return null;      // RLE compressed bitmaps are rare in .docx packages

            byte[] palette = null;
            if (bitCount <= 8)
            {
                if (paletteEntries == 0)
                    paletteEntries = 1 << bitCount;
                var paletteStart = 14 + headerSize;
                if (compression == 3)
                    paletteStart += 12;
                palette = new byte[paletteEntries * 3];
                for (var i = 0; i < paletteEntries; i++)
                {
                    var src = paletteStart + i * paletteEntrySize;
                    if (src + 2 >= data.Length)
                        break;
                    palette[i * 3] = data[src + 2];
                    palette[i * 3 + 1] = data[src + 1];
                    palette[i * 3 + 2] = data[src];
                }
            }

            var rowBytes = ((width * bitCount + 31) / 32) * 4;
            var rgb = new byte[width * height * 3];
            byte[] alpha = null;

            for (var y = 0; y < height; y++)
            {
                var sourceRow = topDown ? y : height - 1 - y;
                var rowStart = pixelOffset + sourceRow * rowBytes;
                if (rowStart < 0 || rowStart + rowBytes > data.Length)
                    continue;

                for (var x = 0; x < width; x++)
                {
                    var target = (y * width + x) * 3;
                    switch (bitCount)
                    {
                        case 1:
                        case 4:
                        case 8:
                        {
                            var bitPos = x * bitCount;
                            var index = (data[rowStart + bitPos / 8] >> (8 - bitCount - bitPos % 8)) & ((1 << bitCount) - 1);
                            if (palette != null && index * 3 + 2 < palette.Length)
                            {
                                rgb[target] = palette[index * 3];
                                rgb[target + 1] = palette[index * 3 + 1];
                                rgb[target + 2] = palette[index * 3 + 2];
                            }
                            break;
                        }
                        case 16:
                        {
                            var value = ReadU16Le(data, rowStart + x * 2);
                            rgb[target] = (byte)(((value >> 10) & 0x1F) * 255 / 31);
                            rgb[target + 1] = (byte)(((value >> 5) & 0x1F) * 255 / 31);
                            rgb[target + 2] = (byte)((value & 0x1F) * 255 / 31);
                            break;
                        }
                        case 24:
                        {
                            var src = rowStart + x * 3;
                            if (src + 2 >= data.Length)
                                break;
                            rgb[target] = data[src + 2];
                            rgb[target + 1] = data[src + 1];
                            rgb[target + 2] = data[src];
                            break;
                        }
                        case 32:
                        {
                            var src = rowStart + x * 4;
                            if (src + 3 >= data.Length)
                                break;
                            rgb[target] = data[src + 2];
                            rgb[target + 1] = data[src + 1];
                            rgb[target + 2] = data[src];
                            if (data[src + 3] != 0xFF)
                            {
                                if (alpha == null)
                                {
                                    alpha = new byte[width * height];
                                    for (var i = 0; i < alpha.Length; i++)
                                        alpha[i] = 255;
                                }
                                alpha[y * width + x] = data[src + 3];
                            }
                            break;
                        }
                        default:
                            return null;
                    }
                }
            }

            // A 32-bit BMP whose alpha channel is entirely zero is opaque in practice.
            if (alpha != null)
            {
                var allZero = true;
                foreach (var a in alpha)
                {
                    if (a != 0)
                    {
                        allZero = false;
                        break;
                    }
                }
                if (allZero)
                    alpha = null;
            }

            return DecodedImage.FromRgb(width, height, rgb, alpha);
        }

        private static uint ReadU32Le(byte[] b, int offset)
        {
            return (uint)(b[offset] | (b[offset + 1] << 8) | (b[offset + 2] << 16) | (b[offset + 3] << 24));
        }

        private static int ReadU16Le(byte[] b, int offset)
        {
            return b[offset] | (b[offset + 1] << 8);
        }
    }

    /// <summary>GIF87a/GIF89a decoder (first frame only).</summary>
    internal static class GifDecoder
    {
        public static DecodedImage Decode(byte[] data)
        {
            if (data == null || data.Length < 13)
                return null;

            var screenWidth = data[6] | (data[7] << 8);
            var screenHeight = data[8] | (data[9] << 8);
            var flags = data[10];
            var pos = 13;

            byte[] globalPalette = null;
            if ((flags & 0x80) != 0)
            {
                var size = 2 << (flags & 0x07);
                globalPalette = ReadPalette(data, pos, size);
                pos += size * 3;
            }

            var transparentIndex = -1;

            while (pos < data.Length)
            {
                var block = data[pos];
                if (block == 0x21)          // extension
                {
                    if (pos + 1 >= data.Length)
                        break;
                    var label = data[pos + 1];
                    var p = pos + 2;
                    if (label == 0xF9 && p < data.Length && data[p] >= 4)
                    {
                        var packed = data[p + 1];
                        if ((packed & 1) != 0)
                            transparentIndex = data[p + 4];
                    }
                    pos = SkipBlocks(data, p);
                }
                else if (block == 0x2C)     // image descriptor
                {
                    if (pos + 10 > data.Length)
                        break;
                    var left = data[pos + 1] | (data[pos + 2] << 8);
                    var top = data[pos + 3] | (data[pos + 4] << 8);
                    var width = data[pos + 5] | (data[pos + 6] << 8);
                    var height = data[pos + 7] | (data[pos + 8] << 8);
                    var localFlags = data[pos + 9];
                    var p = pos + 10;

                    var palette = globalPalette;
                    if ((localFlags & 0x80) != 0)
                    {
                        var size = 2 << (localFlags & 0x07);
                        palette = ReadPalette(data, p, size);
                        p += size * 3;
                    }
                    if (palette == null || width <= 0 || height <= 0)
                        return null;

                    var interlaced = (localFlags & 0x40) != 0;
                    if (p >= data.Length)
                        return null;
                    var minCodeSize = data[p];
                    p++;

                    var indices = LzwDecode(data, ref p, minCodeSize, width * height);
                    if (indices == null)
                        return null;
                    if (interlaced)
                        indices = Deinterlace(indices, width, height);

                    var canvasWidth = Math.Max(screenWidth, left + width);
                    var canvasHeight = Math.Max(screenHeight, top + height);
                    var rgb = new byte[canvasWidth * canvasHeight * 3];
                    byte[] alpha = null;
                    if (transparentIndex >= 0 || left > 0 || top > 0 || width < canvasWidth || height < canvasHeight)
                    {
                        alpha = new byte[canvasWidth * canvasHeight];
                    }

                    for (var y = 0; y < height; y++)
                    {
                        for (var x = 0; x < width; x++)
                        {
                            var index = indices[y * width + x];
                            var targetX = left + x;
                            var targetY = top + y;
                            if (targetX >= canvasWidth || targetY >= canvasHeight)
                                continue;
                            var target = (targetY * canvasWidth + targetX) * 3;
                            if (index * 3 + 2 < palette.Length)
                            {
                                rgb[target] = palette[index * 3];
                                rgb[target + 1] = palette[index * 3 + 1];
                                rgb[target + 2] = palette[index * 3 + 2];
                            }
                            if (alpha != null)
                                alpha[targetY * canvasWidth + targetX] = index == transparentIndex ? (byte)0 : (byte)255;
                        }
                    }

                    return DecodedImage.FromRgb(canvasWidth, canvasHeight, rgb, alpha);
                }
                else
                {
                    break;                  // trailer or unknown block
                }
            }
            return null;
        }

        private static byte[] ReadPalette(byte[] data, int offset, int entries)
        {
            var palette = new byte[entries * 3];
            for (var i = 0; i < entries * 3 && offset + i < data.Length; i++)
                palette[i] = data[offset + i];
            return palette;
        }

        private static int SkipBlocks(byte[] data, int pos)
        {
            while (pos < data.Length)
            {
                var size = data[pos];
                if (size == 0)
                    return pos + 1;
                pos += size + 1;
            }
            return data.Length;
        }

        private static byte[] LzwDecode(byte[] data, ref int pos, int minCodeSize, int pixelCount)
        {
            if (minCodeSize < 2 || minCodeSize > 11)
                return null;

            // Gather the sub-blocks into one buffer.
            var buffer = new List<byte>();
            while (pos < data.Length)
            {
                var size = data[pos];
                pos++;
                if (size == 0)
                    break;
                for (var i = 0; i < size && pos < data.Length; i++, pos++)
                    buffer.Add(data[pos]);
            }

            var clearCode = 1 << minCodeSize;
            var endCode = clearCode + 1;
            var codeSize = minCodeSize + 1;
            var next = endCode + 1;

            var prefix = new int[4096];
            var suffix = new byte[4096];
            for (var i = 0; i < clearCode; i++)
            {
                prefix[i] = -1;
                suffix[i] = (byte)i;
            }

            var output = new byte[pixelCount];
            var outPos = 0;
            var bitPos = 0;
            var previousCode = -1;
            var stack = new byte[4096];
            var totalBits = buffer.Count * 8;

            while (outPos < pixelCount && bitPos + codeSize <= totalBits)
            {
                var code = 0;
                for (var i = 0; i < codeSize; i++)
                {
                    var bit = (buffer[(bitPos + i) / 8] >> ((bitPos + i) % 8)) & 1;
                    code |= bit << i;
                }
                bitPos += codeSize;

                if (code == clearCode)
                {
                    codeSize = minCodeSize + 1;
                    next = endCode + 1;
                    previousCode = -1;
                    continue;
                }
                if (code == endCode)
                    break;

                // Resolve the string for this code onto the stack (reversed).
                var stackTop = 0;
                int current;
                if (code < next)
                {
                    current = code;
                }
                else
                {
                    if (previousCode < 0)
                        break;
                    stack[stackTop++] = FirstChar(prefix, suffix, previousCode);
                    current = previousCode;
                }

                var guard = 0;
                while (current >= clearCode && current < 4096 && guard++ < 4096 && stackTop < stack.Length)
                {
                    stack[stackTop++] = suffix[current];
                    current = prefix[current];
                }
                if (current >= 0 && current < 4096 && stackTop < stack.Length)
                    stack[stackTop++] = suffix[current];

                for (var i = stackTop - 1; i >= 0 && outPos < pixelCount; i--)
                    output[outPos++] = stack[i];

                if (previousCode >= 0 && next < 4096)
                {
                    prefix[next] = previousCode;
                    suffix[next] = stack[stackTop - 1];    // first character of the string just emitted
                    next++;
                    if (next == (1 << codeSize) && codeSize < 12)
                        codeSize++;
                }
                previousCode = code;
            }

            return output;
        }

        private static byte FirstChar(int[] prefix, byte[] suffix, int code)
        {
            var guard = 0;
            while (code >= 0 && prefix[code] >= 0 && guard++ < 4096)
                code = prefix[code];
            return code >= 0 ? suffix[code] : (byte)0;
        }

        private static byte[] Deinterlace(byte[] source, int width, int height)
        {
            var target = new byte[source.Length];
            int[] starts = { 0, 4, 2, 1 };
            int[] steps = { 8, 8, 4, 2 };
            var row = 0;
            for (var pass = 0; pass < 4; pass++)
            {
                for (var y = starts[pass]; y < height; y += steps[pass])
                {
                    if (row * width + width > source.Length)
                        break;
                    Array.Copy(source, row * width, target, y * width, width);
                    row++;
                }
            }
            return target;
        }
    }
}
