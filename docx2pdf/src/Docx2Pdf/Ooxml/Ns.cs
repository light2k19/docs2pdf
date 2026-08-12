using System.Xml.Linq;

namespace Docx2Pdf.Ooxml
{
    /// <summary>Open Packaging / WordprocessingML XML namespaces.</summary>
    internal static class Ns
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        public static readonly XNamespace CT = "http://schemas.openxmlformats.org/package/2006/content-types";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        public static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        public static readonly XNamespace Wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        public static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";

        // Relationship types
        public const string RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const string RelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        public const string RelHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        public const string RelHeader = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        public const string RelFooter = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        public const string RelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const string RelNumbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        public const string RelTheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        public const string RelSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        public const string RelFootnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        public const string RelEndnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    }
}
