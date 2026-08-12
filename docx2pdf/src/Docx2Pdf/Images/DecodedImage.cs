using System;

namespace Docx2Pdf.Images
{
    /// <summary>
    /// A decoded raster image ready to be written as a PDF image XObject.
    /// JPEG data is passed through untouched (DCTDecode); everything else is expanded to 8-bit samples.
    /// </summary>
    internal sealed class DecodedImage
    {
        public int Width;
        public int Height;

        /// <summary>1 = grayscale, 3 = RGB. Ignored for pass-through JPEG.</summary>
        public int Components;

        /// <summary>Interleaved 8-bit samples, Width * Height * Components bytes.</summary>
        public byte[] Samples;

        /// <summary>Optional 8-bit alpha channel, Width * Height bytes.</summary>
        public byte[] Alpha;

        /// <summary>Original JPEG bytes when the image can be embedded without re-encoding.</summary>
        public byte[] JpegData;
        public int JpegComponents;
        public bool JpegAdobeInverted;

        public bool IsJpeg { get { return JpegData != null; } }

        public double AspectRatio
        {
            get { return Height == 0 ? 1 : (double)Width / Height; }
        }

        public static DecodedImage FromRgb(int width, int height, byte[] rgb, byte[] alpha)
        {
            return new DecodedImage { Width = width, Height = height, Components = 3, Samples = rgb, Alpha = alpha };
        }
    }
}
