using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docx2Pdf.Model;

namespace Docx2Pdf.Ooxml
{
    /// <summary>Builds the layout-ready document model from a .docx package.</summary>
    internal sealed class DocumentReader
    {
        private readonly OpcPackage _package;
        private readonly List<string> _warnings;
        private string _documentPart;
        private string _currentPart;          // part whose relationships apply to the nodes being read
        private FormatReader _fmt;
        private StyleSheet _styles;
        private NumberingTable _numbering;
        private readonly Dictionary<string, XDocument> _noteParts = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        private readonly List<Block> _pendingBlocks = new List<Block>();
        private readonly List<FieldState> _fieldStack = new List<FieldState>();
        /// <summary>
        /// Paragraph and run formatting contributed by the table styles currently in scope.
        /// Table styles sit between the document defaults and the paragraph's own style, and are
        /// what makes cell paragraphs single spaced with no space after in most documents.
        /// </summary>
        private readonly List<KeyValuePair<ParagraphFormat, CharacterFormat>> _tableStyleScope =
            new List<KeyValuePair<ParagraphFormat, CharacterFormat>>();
        private string _currentHyperlink;
        private int _noteCounter;

        public StyleSheet Styles { get { return _styles; } }

        private sealed class FieldState
        {
            public StringBuilder Instruction = new StringBuilder();
            public bool InResult;
            public bool SuppressResult;
            public bool IsHyperlink;
        }

        public DocumentReader(OpcPackage package, List<string> warnings)
        {
            _package = package;
            _warnings = warnings ?? new List<string>();
        }

        private void Warn(string message)
        {
            if (!_warnings.Contains(message))
                _warnings.Add(message);
        }

        public WordDocument Read()
        {
            var docRel = _package.FindRelationshipByType(null, Ns.RelOfficeDocument);
            _documentPart = docRel != null
                ? OpcPackage.ResolveTarget(string.Empty, docRel.Target)
                : (_package.HasPart("word/document.xml") ? "word/document.xml" : null);

            if (_documentPart == null || !_package.HasPart(_documentPart))
                throw new InvalidOperationException("The package does not contain a WordprocessingML main document part.");

            _currentPart = _documentPart;

            var themeDoc = ReadRelatedXml(Ns.RelTheme);
            _fmt = new FormatReader(ThemeFonts.Parse(themeDoc));
            _styles = StyleSheet.Parse(ReadRelatedXml(Ns.RelStyles) ?? ReadFallbackXml("word/styles.xml"), _fmt);
            _numbering = NumberingTable.Parse(ReadRelatedXml(Ns.RelNumbering) ?? ReadFallbackXml("word/numbering.xml"), _fmt);

            var result = new WordDocument();
            ReadSettings(result);
            ReadCoreProperties(result);

            var footnotes = ReadRelatedXml(Ns.RelFootnotes) ?? ReadFallbackXml("word/footnotes.xml");
            if (footnotes != null)
                _noteParts["footnote"] = footnotes;
            var endnotes = ReadRelatedXml(Ns.RelEndnotes) ?? ReadFallbackXml("word/endnotes.xml");
            if (endnotes != null)
                _noteParts["endnote"] = endnotes;

            var doc = _package.ReadXml(_documentPart);
            var body = doc.Root == null ? null : doc.Root.Element(Ns.W + "body");
            if (body == null)
                throw new InvalidOperationException("The main document part has no w:body element.");

            var section = new Section();
            var defaults = new SectionProperties();

            foreach (var el in OoxmlUtil.EffectiveElements(body))
            {
                if (el.Name == Ns.W + "p")
                {
                    var pPr = el.Element(Ns.W + "pPr");
                    var sectPr = pPr == null ? null : pPr.Element(Ns.W + "sectPr");

                    var paragraph = ReadParagraph(el);
                    if (paragraph != null)
                        section.Blocks.Add(paragraph);
                    FlushPending(section.Blocks);

                    if (sectPr != null)
                    {
                        section.Properties = ReadSectionProperties(sectPr, defaults);
                        defaults = section.Properties;
                        result.Sections.Add(section);
                        section = new Section();
                    }
                }
                else if (el.Name == Ns.W + "tbl")
                {
                    section.Blocks.Add(ReadTable(el));
                    FlushPending(section.Blocks);
                }
                else if (el.Name == Ns.W + "sectPr")
                {
                    section.Properties = ReadSectionProperties(el, defaults);
                }
                else if (el.Name == Ns.W + "sdt")
                {
                    var content = el.Element(Ns.W + "sdtContent");
                    if (content != null)
                    {
                        foreach (var block in ReadBlocks(content))
                            section.Blocks.Add(block);
                    }
                }
            }

            result.Sections.Add(section);

            // A paragraph-level sectPr at the very end leaves an empty trailing section behind.
            if (result.Sections.Count > 1)
            {
                var last = result.Sections[result.Sections.Count - 1];
                if (last.Blocks.Count == 0 && last.Properties == null)
                    result.Sections.RemoveAt(result.Sections.Count - 1);
            }

            // Every section inherits page setup from the final sectPr when it declared none.
            foreach (var s in result.Sections)
            {
                if (s.Properties == null)
                    s.Properties = defaults;
            }

            AppendNotes(result);
            return result;
        }

        private void FlushPending(List<Block> target)
        {
            if (_pendingBlocks.Count == 0)
                return;
            target.AddRange(_pendingBlocks);
            _pendingBlocks.Clear();
        }

        private XDocument ReadRelatedXml(string relType)
        {
            var rel = _package.FindRelationshipByType(_documentPart, relType);
            if (rel == null || rel.External)
                return null;
            var part = OpcPackage.ResolveTarget(_documentPart, rel.Target);
            return _package.HasPart(part) ? _package.ReadXml(part) : null;
        }

        private XDocument ReadFallbackXml(string part)
        {
            return _package.HasPart(part) ? _package.ReadXml(part) : null;
        }

        private void ReadSettings(WordDocument doc)
        {
            var settings = ReadRelatedXml(Ns.RelSettings) ?? ReadFallbackXml("word/settings.xml");
            if (settings == null || settings.Root == null)
                return;
            doc.EvenOddHeaders = OoxmlUtil.Toggle(settings.Root, Ns.W + "evenAndOddHeaders") == true;
            var tab = settings.Root.Element(Ns.W + "defaultTabStop");
            var val = OoxmlUtil.Dbl(tab, Ns.W + "val");
            if (val.HasValue && val.Value > 0)
                doc.DefaultTabStopPt = OoxmlUtil.TwipsToPoints(val.Value);
        }

        private void ReadCoreProperties(WordDocument doc)
        {
            if (!_package.HasPart("docProps/core.xml"))
                return;
            var core = _package.ReadXml("docProps/core.xml");
            if (core == null || core.Root == null)
                return;
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            doc.Info.Title = (string)core.Root.Element(dc + "title");
            doc.Info.Author = (string)core.Root.Element(dc + "creator");
            doc.Info.Subject = (string)core.Root.Element(dc + "subject");
            doc.Info.Keywords = (string)core.Root.Element(cp + "keywords");
        }

        // ---------------------------------------------------------------- blocks

        private List<Block> ReadBlocks(XElement container)
        {
            var blocks = new List<Block>();
            foreach (var el in OoxmlUtil.EffectiveElements(container))
            {
                if (el.Name == Ns.W + "p")
                {
                    var p = ReadParagraph(el);
                    if (p != null)
                        blocks.Add(p);
                    FlushPending(blocks);
                }
                else if (el.Name == Ns.W + "tbl")
                {
                    blocks.Add(ReadTable(el));
                    FlushPending(blocks);
                }
                else if (el.Name == Ns.W + "sdt")
                {
                    var content = el.Element(Ns.W + "sdtContent");
                    if (content != null)
                        blocks.AddRange(ReadBlocks(content));
                }
            }
            return blocks;
        }

        private Paragraph ReadParagraph(XElement el)
        {
            var p = new Paragraph();
            var pPr = el.Element(Ns.W + "pPr");

            p.StyleId = OoxmlUtil.ChildVal(pPr, Ns.W + "pStyle") ?? _styles.DefaultParagraphStyleId;

            var format = _styles.DefaultParagraphFormat.Clone();
            var runDefaults = _styles.DefaultCharacterFormat.Clone();

            // Enclosing table styles apply before the paragraph's own style (outermost table first).
            foreach (var scope in _tableStyleScope)
            {
                format.ApplyOver(scope.Key);
                runDefaults.ApplyOver(scope.Value);
            }

            format.ApplyOver(_styles.ParagraphFormatFor(p.StyleId));
            runDefaults.ApplyOver(_styles.CharacterFormatFor(p.StyleId));

            // Numbering, either direct or inherited from the paragraph style.
            string numId = null;
            var ilvl = 0;
            var numPr = pPr == null ? null : pPr.Element(Ns.W + "numPr");
            if (numPr != null)
            {
                numId = OoxmlUtil.ChildVal(numPr, Ns.W + "numId");
                var lvlText = OoxmlUtil.ChildVal(numPr, Ns.W + "ilvl");
                int lvl;
                if (lvlText != null && int.TryParse(lvlText, out lvl))
                    ilvl = lvl;
            }
            if (numId == null && p.StyleId != null)
            {
                string styleNumId;
                int styleLevel;
                if (_styles.TryGetStyleNumbering(p.StyleId, out styleNumId, out styleLevel))
                {
                    numId = styleNumId;
                    ilvl = styleLevel;

                    // A level bound to this style through w:lvl/w:pStyle is authoritative.
                    int linkedLevel;
                    if (_numbering.TryGetLevelForStyle(numId, p.StyleId, out linkedLevel))
                        ilvl = linkedLevel;
                }
            }

            var directFormat = _fmt.ReadParagraphFormat(pPr);
            format.ApplyOver(directFormat);

            // Direct run properties on the paragraph mark style the pilcrow only: they set the
            // height of an empty paragraph but never restyle the paragraph's text runs.
            var markFormat = runDefaults;
            if (pPr != null)
            {
                var markRPr = pPr.Element(Ns.W + "rPr");
                if (markRPr != null)
                {
                    markFormat = runDefaults.Clone();
                    var rStyle = OoxmlUtil.ChildVal(markRPr, Ns.W + "rStyle");
                    if (rStyle != null)
                        markFormat.ApplyOver(_styles.CharacterFormatFor(rStyle));
                    markFormat.ApplyOver(_fmt.ReadRunFormat(markRPr));
                }
            }

            p.Format = format;
            p.RunDefaults = runDefaults;
            p.MarkFormat = markFormat;
            p.HeadingLevel = _styles.HeadingLevelOf(p.StyleId);
            if (p.HeadingLevel == 0 && format.OutlineLevel.HasValue && format.OutlineLevel.Value <= 8)
                p.HeadingLevel = format.OutlineLevel.Value + 1;

            if (numId != null && numId != "0" && _numbering.HasInstance(numId))
                ApplyNumbering(p, numId, ilvl, runDefaults, directFormat);

            ReadInlines(el, p, runDefaults);
            return p;
        }

        private void ApplyNumbering(Paragraph p, string numId, int ilvl, CharacterFormat runDefaults,
                                    ParagraphFormat directFormat)
        {
            NumberingLevel level;
            var label = _numbering.NextLabel(numId, ilvl, out level);
            if (label == null || level == null)
                return;

            // The list symbol takes the paragraph mark's formatting (w:pPr/w:rPr sizes the
            // number/bullet), overridden by the level's own rPr.
            var labelFormat = (p.MarkFormat ?? runDefaults).Clone();
            if (level.RunFormat != null)
            {
                var lf = level.RunFormat.Clone();
                labelFormat.ApplyOver(lf);
            }
            if (level.Format == NumberFormat.Bullet && !string.IsNullOrEmpty(level.BulletFont))
                labelFormat.FontFamily = level.BulletFont;

            p.ListLabel = label;
            p.ListNumId = numId;
            p.ListLabelFormat = labelFormat;
            p.ListLabelFollowedByTab = level.Suffix != "nothing";
            if (level.Suffix == "space")
                p.ListLabelFollowedByTab = false;

            // Word's indent precedence for numbered paragraphs: direct w:ind on the paragraph
            // wins; otherwise the numbering level's w:ind replaces any style-derived indent.
            var indentLeft = p.Format.IndentLeftPt;
            var firstLine = p.Format.IndentFirstLinePt;
            if (level.HasIndent)
            {
                indentLeft = directFormat != null && directFormat.IndentLeftPt.HasValue
                    ? directFormat.IndentLeftPt
                    : (double?)level.IndentLeftPt;
                firstLine = directFormat != null && directFormat.IndentFirstLinePt.HasValue
                    ? directFormat.IndentFirstLinePt
                    : (level.HangingPt > 0 ? -level.HangingPt : (double?)0);
            }
            p.Format.IndentLeftPt = indentLeft;
            p.Format.IndentFirstLinePt = firstLine;
            p.ListTextIndentPt = indentLeft ?? 0;
            p.ListLabelIndentPt = (indentLeft ?? 0) + (firstLine ?? 0);
            if (level.Suffix == "space")
                p.ListLabel = label + " ";
        }

        private void ReadInlines(XElement container, Paragraph p, CharacterFormat inherited)
        {
            foreach (var el in OoxmlUtil.EffectiveElements(container))
            {
                var name = el.Name;
                if (name == Ns.W + "pPr")
                    continue;

                if (name == Ns.W + "r")
                {
                    ReadRun(el, p, inherited);
                }
                else if (name == Ns.W + "hyperlink")
                {
                    var previous = _currentHyperlink;
                    string anchor = OoxmlUtil.Str(el, Ns.W + "anchor");
                    var relId = OoxmlUtil.Str(el, Ns.R + "id");
                    string url = null;
                    if (!string.IsNullOrEmpty(relId))
                    {
                        var rel = _package.GetRelationship(_currentPart, relId);
                        if (rel != null)
                            url = rel.External ? rel.Target : null;
                    }
                    _currentHyperlink = url;
                    var startIndex = p.Inlines.Count;
                    ReadInlines(el, p, inherited);
                    if (!string.IsNullOrEmpty(anchor) && url == null)
                    {
                        for (var i = startIndex; i < p.Inlines.Count; i++)
                            p.Inlines[i].LinkAnchor = anchor;
                    }
                    _currentHyperlink = previous;
                }
                else if (name == Ns.W + "ins" || name == Ns.W + "smartTag" || name == Ns.W + "sdtContent"
                         || name == Ns.W + "bdo" || name == Ns.W + "dir")
                {
                    ReadInlines(el, p, inherited);
                }
                else if (name == Ns.W + "sdt")
                {
                    var content = el.Element(Ns.W + "sdtContent");
                    if (content != null)
                        ReadInlines(content, p, inherited);
                }
                else if (name == Ns.W + "del" || name == Ns.W + "moveFrom")
                {
                    // Rejected/moved-away content is not rendered.
                }
                else if (name == Ns.W + "moveTo")
                {
                    ReadInlines(el, p, inherited);
                }
                else if (name == Ns.W + "bookmarkStart")
                {
                    var bookmarkName = OoxmlUtil.Str(el, Ns.W + "name");
                    if (!string.IsNullOrEmpty(bookmarkName) && !bookmarkName.StartsWith("_GoBack", StringComparison.Ordinal))
                        p.Inlines.Add(new BookmarkInline { Name = bookmarkName, Format = inherited });
                }
                else if (name == Ns.W + "fldSimple")
                {
                    var instr = OoxmlUtil.Str(el, Ns.W + "instr") ?? string.Empty;
                    var kind = ClassifyField(instr);
                    if (kind != FieldKind.None)
                    {
                        p.Inlines.Add(new FieldInline { Kind = kind, Format = inherited.Clone(), LinkUrl = _currentHyperlink });
                    }
                    else
                    {
                        var url = HyperlinkTargetOf(instr);
                        var previous = _currentHyperlink;
                        if (url != null)
                            _currentHyperlink = url;
                        ReadInlines(el, p, inherited);
                        _currentHyperlink = previous;
                    }
                }
                else if (name == Ns.W + "subDoc")
                {
                    Warn("Subdocument references are not expanded.");
                }
            }
        }

        private void ReadRun(XElement run, Paragraph p, CharacterFormat inherited)
        {
            var format = inherited.Clone();
            var rPr = run.Element(Ns.W + "rPr");
            if (rPr != null)
            {
                var rStyle = OoxmlUtil.ChildVal(rPr, Ns.W + "rStyle");
                if (rStyle != null)
                    format.ApplyOver(_styles.CharacterFormatFor(rStyle));
                format.ApplyOver(_fmt.ReadRunFormat(rPr));
            }

            var suppressed = _fieldStack.Count > 0 && _fieldStack[_fieldStack.Count - 1].InResult
                             && _fieldStack[_fieldStack.Count - 1].SuppressResult;

            foreach (var el in OoxmlUtil.EffectiveElements(run))
            {
                var name = el.Name;
                if (name == Ns.W + "rPr")
                    continue;

                if (name == Ns.W + "t")
                {
                    if (!suppressed)
                        AddText(p, el.Value, format);
                }
                else if (name == Ns.W + "delText")
                {
                    // not rendered
                }
                else if (name == Ns.W + "instrText")
                {
                    if (_fieldStack.Count > 0 && !_fieldStack[_fieldStack.Count - 1].InResult)
                        _fieldStack[_fieldStack.Count - 1].Instruction.Append(el.Value);
                }
                else if (name == Ns.W + "fldChar")
                {
                    HandleFieldChar(el, p, format);
                }
                else if (name == Ns.W + "tab")
                {
                    if (!suppressed)
                        p.Inlines.Add(new TabInline { Format = format, LinkUrl = _currentHyperlink });
                }
                else if (name == Ns.W + "br")
                {
                    var type = (OoxmlUtil.Str(el, Ns.W + "type") ?? "textWrapping").ToLowerInvariant();
                    var kind = type == "page" ? BreakKind.Page : (type == "column" ? BreakKind.Column : BreakKind.Line);
                    p.Inlines.Add(new BreakInline { Kind = kind, Format = format });
                }
                else if (name == Ns.W + "cr")
                {
                    p.Inlines.Add(new BreakInline { Kind = BreakKind.Line, Format = format });
                }
                else if (name == Ns.W + "noBreakHyphen")
                {
                    if (!suppressed)
                        AddText(p, "‑", format);
                }
                else if (name == Ns.W + "softHyphen")
                {
                    // Rendered only when the line breaks there; ignored.
                }
                else if (name == Ns.W + "sym")
                {
                    if (!suppressed)
                        AddSymbol(el, p, format);
                }
                else if (name == Ns.W + "drawing" || name == Ns.W + "pict" || name == Ns.W + "object")
                {
                    if (!suppressed)
                        ReadGraphics(el, p, format);
                }
                else if (name == Ns.W + "footnoteReference" || name == Ns.W + "endnoteReference")
                {
                    ReadNoteReference(el, p, format, name.LocalName == "footnoteReference" ? "footnote" : "endnote");
                }
                else if (name == Ns.W + "footnoteRef" || name == Ns.W + "endnoteRef")
                {
                    // Placeholder for the note's own number; the note body already carries it.
                }
                else if (name == Ns.W + "ptab")
                {
                    p.Inlines.Add(new TabInline { Format = format });
                }
            }
        }

        private void AddText(Paragraph p, string text, CharacterFormat format)
        {
            if (string.IsNullOrEmpty(text))
                return;
            if (format.Hidden == true)
                return;
            p.Inlines.Add(new TextInline(text, format) { LinkUrl = _currentHyperlink });
        }

        private void AddSymbol(XElement el, Paragraph p, CharacterFormat format)
        {
            var charCode = OoxmlUtil.Str(el, Ns.W + "char");
            var font = OoxmlUtil.Str(el, Ns.W + "font");
            if (string.IsNullOrEmpty(charCode))
                return;
            int code;
            if (!int.TryParse(charCode, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                return;

            var symbolFormat = format.Clone();
            if (!string.IsNullOrEmpty(font))
                symbolFormat.FontFamily = font;
            AddText(p, char.ConvertFromUtf32(code), symbolFormat);
        }

        private void HandleFieldChar(XElement el, Paragraph p, CharacterFormat format)
        {
            var type = (OoxmlUtil.Str(el, Ns.W + "fldCharType") ?? string.Empty).ToLowerInvariant();
            if (type == "begin")
            {
                _fieldStack.Add(new FieldState());
            }
            else if (type == "separate")
            {
                if (_fieldStack.Count == 0)
                    return;
                var state = _fieldStack[_fieldStack.Count - 1];
                state.InResult = true;
                var instruction = state.Instruction.ToString();
                var kind = ClassifyField(instruction);
                if (kind != FieldKind.None)
                {
                    state.SuppressResult = true;
                    p.Inlines.Add(new FieldInline { Kind = kind, Format = format.Clone(), LinkUrl = _currentHyperlink });
                }
                else
                {
                    var url = HyperlinkTargetOf(instruction);
                    if (url != null)
                    {
                        state.IsHyperlink = true;
                        _currentHyperlink = url;
                    }
                }
            }
            else if (type == "end")
            {
                if (_fieldStack.Count == 0)
                    return;
                var state = _fieldStack[_fieldStack.Count - 1];
                _fieldStack.RemoveAt(_fieldStack.Count - 1);
                if (state.IsHyperlink)
                    _currentHyperlink = null;
            }
        }

        private static FieldKind ClassifyField(string instruction)
        {
            if (string.IsNullOrEmpty(instruction))
                return FieldKind.None;
            var trimmed = instruction.Trim();
            var space = trimmed.IndexOf(' ');
            var keyword = (space > 0 ? trimmed.Substring(0, space) : trimmed).ToUpperInvariant();
            switch (keyword)
            {
                case "PAGE": return FieldKind.Page;
                case "NUMPAGES": return FieldKind.NumPages;
                case "SECTIONPAGES": return FieldKind.SectionPages;
                default: return FieldKind.None;
            }
        }

        private static string HyperlinkTargetOf(string instruction)
        {
            if (string.IsNullOrEmpty(instruction))
                return null;
            var trimmed = instruction.Trim();
            if (!trimmed.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                return null;
            var rest = trimmed.Substring("HYPERLINK".Length).Trim();
            if (rest.StartsWith("\"", StringComparison.Ordinal))
            {
                var end = rest.IndexOf('"', 1);
                if (end > 1)
                    return rest.Substring(1, end - 1);
            }
            var token = rest.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return token.Length > 0 && !token[0].StartsWith("\\", StringComparison.Ordinal) ? token[0] : null;
        }

        private void ReadNoteReference(XElement el, Paragraph p, CharacterFormat format, string kind)
        {
            var id = OoxmlUtil.Str(el, Ns.W + "id");
            if (id == null)
                return;

            XDocument part;
            if (!_noteParts.TryGetValue(kind, out part) || part.Root == null)
                return;

            var noteEl = part.Root.Elements(Ns.W + (kind == "footnote" ? "footnote" : "endnote"))
                .FirstOrDefault(n => OoxmlUtil.Str(n, Ns.W + "id") == id);
            if (noteEl == null)
                return;
            var noteType = (OoxmlUtil.Str(noteEl, Ns.W + "type") ?? string.Empty).ToLowerInvariant();
            if (noteType == "separator" || noteType == "continuationseparator" || noteType == "continuationnotice")
                return;

            _noteCounter++;
            var marker = _noteCounter.ToString(CultureInfo.InvariantCulture);
            var markerFormat = format.Clone();
            markerFormat.VertAlign = VerticalTextAlignment.Superscript;
            p.Inlines.Add(new TextInline(marker, markerFormat));

            var blocks = ReadBlocks(noteEl);
            if (blocks.Count > 0)
            {
                // Prefix the note body with its marker.
                var first = blocks[0] as Paragraph;
                if (first != null)
                {
                    var numberFormat = first.RunDefaults.Clone();
                    numberFormat.VertAlign = VerticalTextAlignment.Superscript;
                    first.Inlines.Insert(0, new TextInline(marker + " ", numberFormat));
                }
                _pendingNotes.Add(new KeyValuePair<string, List<Block>>(marker, blocks));
            }
        }

        private readonly List<KeyValuePair<string, List<Block>>> _pendingNotes = new List<KeyValuePair<string, List<Block>>>();

        private void AppendNotes(WordDocument doc)
        {
            foreach (var note in _pendingNotes)
                doc.Notes.Add(note);
        }

        // ---------------------------------------------------------------- graphics

        private void ReadGraphics(XElement el, Paragraph p, CharacterFormat format)
        {
            if (el.Name == Ns.W + "drawing")
            {
                foreach (var anchor in el.Elements())
                    ReadDrawingAnchor(anchor, p, format);
            }
            else
            {
                ReadVml(el, p, format);
            }
        }

        private void ReadDrawingAnchor(XElement anchor, Paragraph p, CharacterFormat format)
        {
            double widthPt = 0, heightPt = 0;
            var extent = anchor.Element(Ns.Wp + "extent");
            if (extent != null)
            {
                widthPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(extent, "cx") ?? 0);
                heightPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(extent, "cy") ?? 0);
            }

            var description = OoxmlUtil.Str(anchor.Descendants(Ns.Wp + "docPr").FirstOrDefault(), "descr");

            // Floating objects keep their own position instead of flowing with the text.
            if (anchor.Name == Ns.Wp + "anchor")
            {
                ReadFloatingAnchor(anchor, p, format, widthPt, heightPt, description);
                return;
            }

            // A wordprocessingGroup (a screenshot with annotation shapes in a nested
            // coordinate space) must be handled before the text-box path below: the
            // group's first descendant text box is usually an empty annotation label,
            // and taking it swallowed the whole picture (MyLesen p51: the receipt
            // screenshot vanished and the next script's title block slid up a page).
            var group = anchor.Descendants(Ns.Wpg + "wgp").FirstOrDefault();
            if (group != null)
            {
                ReadGroupShape(group, p, format, widthPt, heightPt, description);
                return;
            }

            // Inline text boxes become follow-on blocks so their content is not lost.
            var textBox = anchor.Descendants(Ns.W + "txbxContent").FirstOrDefault();
            if (textBox != null)
            {
                _pendingBlocks.AddRange(ReadBlocks(textBox));
                return;
            }

            var blip = anchor.Descendants(Ns.A + "blip").FirstOrDefault();
            if (blip == null)
            {
                var shapeName = anchor.Descendants(Ns.A + "prstGeom").FirstOrDefault();
                if (shapeName != null)
                    Warn("Vector shapes and charts are rendered as blank space.");
                return;
            }

            var relId = OoxmlUtil.Str(blip, Ns.R + "embed") ?? OoxmlUtil.Str(blip, Ns.R + "link");
            var rotation = 0.0;
            var xfrm = anchor.Descendants(Ns.A + "xfrm").FirstOrDefault();
            if (xfrm != null)
            {
                var rot = OoxmlUtil.Dbl(xfrm, "rot");
                if (rot.HasValue)
                    rotation = rot.Value / 60000.0;
                if (widthPt <= 0 || heightPt <= 0)
                {
                    var ext = xfrm.Element(Ns.A + "ext");
                    if (ext != null)
                    {
                        widthPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(ext, "cx") ?? 0);
                        heightPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(ext, "cy") ?? 0);
                    }
                }
            }

            var image = CreateImage(relId, widthPt, heightPt, rotation, description, format);
            if (image == null)
                return;
            ReadCrop(anchor, image);
            ReadPictureFrame(anchor, image);

            // Effects such as drop shadows occupy extra layout space around the picture.
            var effectExtent = anchor.Element(Ns.Wp + "effectExtent");
            if (effectExtent != null)
            {
                image.EffectLeftPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(effectExtent, "l") ?? 0);
                image.EffectTopPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(effectExtent, "t") ?? 0);
                image.EffectRightPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(effectExtent, "r") ?? 0);
                image.EffectBottomPt = OoxmlUtil.EmuToPoints(OoxmlUtil.Dbl(effectExtent, "b") ?? 0);
            }
            p.Inlines.Add(image);
        }

        private sealed class GroupPicture
        {
            public string RelId;
            public double X, Y, Width, Height;
            public XElement Element;
        }

        /// <summary>
        /// Renders a wpg:wgp group: the (largest) picture plus simple rectangle shapes
        /// drawn over it. Children live in the group's chOff/chExt coordinate space and
        /// map through nested group transforms onto the drawing's extent (MyLesen p51:
        /// a 193x169pt receipt screenshot with a red highlight rectangle).
        /// </summary>
        private void ReadGroupShape(XElement group, Paragraph p, CharacterFormat format,
                                    double widthPt, double heightPt, string description)
        {
            if (widthPt <= 0 || heightPt <= 0)
                return;

            var xfrm = GroupXfrm(group);
            var extX = EmuAttr(xfrm, "ext", "cx");
            var extY = EmuAttr(xfrm, "ext", "cy");
            if (extX <= 0 || extY <= 0)
                return;

            // Maps the group's parent space onto the drawing canvas (points): the group's
            // outer rectangle covers (0,0)..(widthPt,heightPt).
            var sx = widthPt / extX;
            var sy = heightPt / extY;
            var dx = -EmuAttr(xfrm, "off", "x") * sx;
            var dy = -EmuAttr(xfrm, "off", "y") * sy;

            var pictures = new List<GroupPicture>();
            var overlays = new List<ImageOverlay>();
            WalkGroup(group, sx, sy, dx, dy, pictures, overlays);
            if (pictures.Count == 0)
            {
                Warn("A group shape without a picture is rendered as blank space.");
                return;
            }

            var main = pictures[0];
            foreach (var candidate in pictures)
            {
                if (candidate.Width * candidate.Height > main.Width * main.Height)
                    main = candidate;
            }
            if (pictures.Count > 1)
                Warn("Only the largest picture of a group shape is rendered.");

            var image = CreateImage(main.RelId, main.Width, main.Height, 0, description, format);
            if (image == null)
                return;
            ReadCrop(main.Element, image);
            ReadPictureFrame(main.Element, image);

            // The picture sits inside the group's box at its mapped offset; the effect
            // margins carry that placement so the layout box equals the group extent.
            image.EffectLeftPt = Math.Max(0, main.X);
            image.EffectTopPt = Math.Max(0, main.Y);
            image.EffectRightPt = Math.Max(0, widthPt - main.X - main.Width);
            image.EffectBottomPt = Math.Max(0, heightPt - main.Y - main.Height);
            if (overlays.Count > 0)
                image.Overlays = overlays;
            p.Inlines.Add(image);
        }

        /// <summary>
        /// Walks a group's children with (sx, sy, dx, dy) mapping the GROUP'S PARENT
        /// space to canvas points; children first pass through the group's own
        /// chOff/chExt -> off/ext transform.
        /// </summary>
        private static void WalkGroup(XElement group, double sx, double sy, double dx, double dy,
                                      List<GroupPicture> pictures, List<ImageOverlay> overlays)
        {
            var xfrm = GroupXfrm(group);
            var chExtX = EmuAttr(xfrm, "chExt", "cx");
            var chExtY = EmuAttr(xfrm, "chExt", "cy");
            if (chExtX <= 0 || chExtY <= 0)
                return;
            var extX = EmuAttr(xfrm, "ext", "cx");
            var extY = EmuAttr(xfrm, "ext", "cy");
            var offX = EmuAttr(xfrm, "off", "x");
            var offY = EmuAttr(xfrm, "off", "y");
            var chOffX = EmuAttr(xfrm, "chOff", "x");
            var chOffY = EmuAttr(xfrm, "chOff", "y");

            var csx = sx * extX / chExtX;
            var csy = sy * extY / chExtY;
            var cdx = dx + sx * (offX - chOffX * extX / chExtX);
            var cdy = dy + sy * (offY - chOffY * extY / chExtY);

            foreach (var child in group.Elements())
            {
                if (child.Name == Ns.Wpg + "grpSp" || child.Name == Ns.Wpg + "wgp")
                {
                    WalkGroup(child, csx, csy, cdx, cdy, pictures, overlays);
                }
                else if (child.Name == Ns.Pic + "pic")
                {
                    var spXfrm = child.Descendants(Ns.A + "xfrm").FirstOrDefault();
                    var blip = child.Descendants(Ns.A + "blip").FirstOrDefault();
                    if (spXfrm == null || blip == null)
                        continue;
                    pictures.Add(new GroupPicture
                    {
                        RelId = OoxmlUtil.Str(blip, Ns.R + "embed") ?? OoxmlUtil.Str(blip, Ns.R + "link"),
                        X = cdx + EmuAttr(spXfrm, "off", "x") * csx,
                        Y = cdy + EmuAttr(spXfrm, "off", "y") * csy,
                        Width = EmuAttr(spXfrm, "ext", "cx") * csx,
                        Height = EmuAttr(spXfrm, "ext", "cy") * csy,
                        Element = child,
                    });
                }
                else if (child.Name == Ns.Wps + "wsp")
                {
                    var spPr = child.Element(Ns.Wps + "spPr");
                    if (spPr == null || spPr.Element(Ns.A + "prstGeom") == null)
                        continue;
                    var spXfrm = spPr.Element(Ns.A + "xfrm");
                    if (spXfrm == null)
                        continue;

                    var overlay = new ImageOverlay
                    {
                        X = cdx + EmuAttr(spXfrm, "off", "x") * csx,
                        Y = cdy + EmuAttr(spXfrm, "off", "y") * csy,
                        Width = EmuAttr(spXfrm, "ext", "cx") * csx,
                        Height = EmuAttr(spXfrm, "ext", "cy") * csy,
                    };
                    var ln = spPr.Element(Ns.A + "ln");
                    if (ln != null)
                    {
                        var color = ShapeColor(ln.Element(Ns.A + "solidFill"));
                        if (color.HasValue)
                        {
                            overlay.OutlineColor = color;
                            var w = OoxmlUtil.Dbl(ln, "w");
                            overlay.OutlineWidthPt = w.HasValue ? Math.Max(0.5, w.Value / 12700.0) : 1;
                        }
                    }
                    overlay.FillColor = ShapeColor(spPr.Element(Ns.A + "solidFill"));
                    // Unstyled rectangles (canvas placeholders, empty label boxes) draw nothing.
                    if (overlay.OutlineColor.HasValue || overlay.FillColor.HasValue)
                        overlays.Add(overlay);
                }
            }
        }

        private static XElement GroupXfrm(XElement group)
        {
            var props = group.Element(Ns.Wpg + "grpSpPr");
            return props == null ? null : props.Element(Ns.A + "xfrm");
        }

        private static double EmuAttr(XElement xfrm, string child, string attribute)
        {
            var el = xfrm == null ? null : xfrm.Element(Ns.A + child);
            var value = el == null ? null : OoxmlUtil.Dbl(el, attribute);
            return value ?? 0;
        }

        /// <summary>Reads a wp:anchor into a floating frame with its own page position.</summary>
        private void ReadFloatingAnchor(XElement anchor, Paragraph p, CharacterFormat format,
                                        double widthPt, double heightPt, string description)
        {
            var frame = new AnchoredInline
            {
                WidthPt = widthPt,
                HeightPt = heightPt,
                Format = format,
                BehindDoc = OoxmlUtil.Str(anchor, "behindDoc") == "1",
            };

            var positionH = anchor.Element(Ns.Wp + "positionH");
            if (positionH != null)
            {
                frame.HorizontalFrom = ParseRelativeFrom(OoxmlUtil.Str(positionH, "relativeFrom"), AnchorRelativeFrom.Column);
                var offset = positionH.Element(Ns.Wp + "posOffset");
                if (offset != null)
                {
                    double emu;
                    if (double.TryParse(offset.Value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out emu))
                        frame.HorizontalOffsetPt = OoxmlUtil.EmuToPoints(emu);
                }
                else
                {
                    var align = positionH.Element(Ns.Wp + "align");
                    if (align != null)
                        frame.HorizontalAlign = align.Value.Trim().ToLowerInvariant();
                }
            }

            var positionV = anchor.Element(Ns.Wp + "positionV");
            if (positionV != null)
            {
                frame.VerticalFrom = ParseRelativeFrom(OoxmlUtil.Str(positionV, "relativeFrom"), AnchorRelativeFrom.Paragraph);
                var offset = positionV.Element(Ns.Wp + "posOffset");
                if (offset != null)
                {
                    double emu;
                    if (double.TryParse(offset.Value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out emu))
                        frame.VerticalOffsetPt = OoxmlUtil.EmuToPoints(emu);
                }
                else
                {
                    var align = positionV.Element(Ns.Wp + "align");
                    if (align != null)
                        frame.VerticalAlign = align.Value.Trim().ToLowerInvariant();
                }
            }

            var textBox = anchor.Descendants(Ns.W + "txbxContent").FirstOrDefault();
            if (textBox != null)
            {
                frame.Blocks = ReadBlocks(textBox);
            }
            else
            {
                var blip = anchor.Descendants(Ns.A + "blip").FirstOrDefault();
                if (blip == null)
                {
                    // A bare preset-geometry shape (e.g. a highlight rectangle drawn over a
                    // screenshot) renders as its outline and fill; anything richer is skipped.
                    if (ReadShapeStyle(anchor, frame))
                    {
                        p.Inlines.Add(frame);
                        return;
                    }
                    Warn("Vector shapes and charts are rendered as blank space.");
                    return;
                }

                var relId = OoxmlUtil.Str(blip, Ns.R + "embed") ?? OoxmlUtil.Str(blip, Ns.R + "link");
                var rotation = 0.0;
                var xfrm = anchor.Descendants(Ns.A + "xfrm").FirstOrDefault();
                var rot = OoxmlUtil.Dbl(xfrm, "rot");
                if (rot.HasValue)
                    rotation = rot.Value / 60000.0;

                var image = CreateImage(relId, widthPt, heightPt, rotation, description, format);
                if (image == null)
                    return;
                ReadCrop(anchor, image);

                var holder = new Paragraph
                {
                    Format = _styles.DefaultParagraphFormat.Clone(),
                    RunDefaults = _styles.DefaultCharacterFormat.Clone(),
                };
                holder.Format.SpaceBeforePt = 0;
                holder.Format.SpaceAfterPt = 0;
                holder.Format.IndentLeftPt = 0;
                holder.Format.IndentRightPt = 0;
                holder.Format.IndentFirstLinePt = 0;
                holder.Format.LineSpacing = 1;
                holder.Format.LineSpacingRule = LineSpacingRule.Exact;
                holder.Format.LineSpacing = Math.Max(1, heightPt);
                holder.Inlines.Add(image);
                frame.Blocks.Add(holder);
            }

            if (frame.Blocks.Count == 0)
                return;
            p.Inlines.Add(frame);
        }

        /// <summary>
        /// Reads outline and fill of a simple preset-geometry shape into the frame.
        /// Returns false when the drawing is not a shape this renderer can represent.
        /// </summary>
        private static bool ReadShapeStyle(XElement anchor, AnchoredInline frame)
        {
            var spPr = anchor.Descendants(Ns.Wps + "spPr").FirstOrDefault();
            if (spPr == null || spPr.Element(Ns.A + "prstGeom") == null)
                return false;

            var any = false;
            var ln = spPr.Element(Ns.A + "ln");
            if (ln != null)
            {
                var color = ShapeColor(ln.Element(Ns.A + "solidFill"));
                if (color.HasValue)
                {
                    frame.OutlineColor = color;
                    var w = OoxmlUtil.Dbl(ln, "w");
                    frame.OutlineWidthPt = w.HasValue && w.Value > 0 ? w.Value / 12700.0 : 1;
                    any = true;
                }
            }
            var fill = ShapeColor(spPr.Element(Ns.A + "solidFill"));
            if (fill.HasValue)
            {
                frame.FillColor = fill;
                any = true;
            }
            return any;
        }

        private static uint? ShapeColor(XElement solidFill)
        {
            if (solidFill == null)
                return null;
            var srgb = solidFill.Element(Ns.A + "srgbClr");
            var hex = srgb == null ? null : OoxmlUtil.Str(srgb, "val");
            uint value;
            if (hex != null && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value))
                return value;
            return null;
        }

        private static AnchorRelativeFrom ParseRelativeFrom(string value, AnchorRelativeFrom fallback)
        {
            if (string.IsNullOrEmpty(value))
                return fallback;
            switch (value.ToLowerInvariant())
            {
                case "page": return AnchorRelativeFrom.Page;
                case "margin": return AnchorRelativeFrom.Margin;
                case "leftmargin": return AnchorRelativeFrom.LeftMargin;
                case "rightmargin": return AnchorRelativeFrom.RightMargin;
                case "topmargin": return AnchorRelativeFrom.TopMargin;
                case "bottommargin": return AnchorRelativeFrom.BottomMargin;
                case "insidemargin": return AnchorRelativeFrom.InsideMargin;
                case "outsidemargin": return AnchorRelativeFrom.OutsideMargin;
                case "column": return AnchorRelativeFrom.Column;
                case "character": return AnchorRelativeFrom.Character;
                case "paragraph": return AnchorRelativeFrom.Paragraph;
                case "line": return AnchorRelativeFrom.Line;
                default: return fallback;
            }
        }

        private void ReadVml(XElement el, Paragraph p, CharacterFormat format)
        {
            var textBox = el.Descendants(Ns.W + "txbxContent").FirstOrDefault();
            if (textBox != null)
            {
                _pendingBlocks.AddRange(ReadBlocks(textBox));
                return;
            }

            var imageData = el.Descendants(Ns.V + "imagedata").FirstOrDefault();
            if (imageData == null)
                return;
            var relId = OoxmlUtil.Str(imageData, Ns.R + "id") ?? OoxmlUtil.Str(imageData, Ns.R + "href");

            double widthPt = 0, heightPt = 0;
            var shape = imageData.Parent;
            var style = shape == null ? null : OoxmlUtil.Str(shape, "style");
            if (!string.IsNullOrEmpty(style))
            {
                widthPt = ParseCssLength(style, "width");
                heightPt = ParseCssLength(style, "height");
            }
            AddImage(p, format, relId, widthPt, heightPt, 0, OoxmlUtil.Str(shape, "alt"));
        }

        private static double ParseCssLength(string style, string property)
        {
            foreach (var part in style.Split(';'))
            {
                var idx = part.IndexOf(':');
                if (idx <= 0)
                    continue;
                var key = part.Substring(0, idx).Trim();
                if (!string.Equals(key, property, StringComparison.OrdinalIgnoreCase))
                    continue;
                var value = part.Substring(idx + 1).Trim();
                var unit = "pt";
                foreach (var u in new[] { "pt", "px", "in", "cm", "mm", "pc" })
                {
                    if (value.EndsWith(u, StringComparison.OrdinalIgnoreCase))
                    {
                        unit = u;
                        value = value.Substring(0, value.Length - u.Length);
                        break;
                    }
                }
                double number;
                if (!double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                    return 0;
                switch (unit.ToLowerInvariant())
                {
                    case "px": return number * 0.75;
                    case "in": return number * 72.0;
                    case "cm": return number * 72.0 / 2.54;
                    case "mm": return number * 72.0 / 25.4;
                    case "pc": return number * 12.0;
                    default: return number;
                }
            }
            return 0;
        }

        private void AddImage(Paragraph p, CharacterFormat format, string relId,
                              double widthPt, double heightPt, double rotation, string description)
        {
            var image = CreateImage(relId, widthPt, heightPt, rotation, description, format);
            if (image != null)
                p.Inlines.Add(image);
        }

        /// <summary>
        /// Reads a picture frame (pic:spPr/a:ln): Word centres the stroke on the extent
        /// boundary, so the visible photo is inset by half the stroke width.
        /// </summary>
        private static void ReadPictureFrame(XElement scope, ImageInline image)
        {
            var spPr = scope.Descendants(Ns.Pic + "spPr").FirstOrDefault();
            if (spPr == null)
                return;
            var ln = spPr.Element(Ns.A + "ln");
            if (ln == null)
                return;
            var w = OoxmlUtil.Dbl(ln, "w");
            if (!w.HasValue || w.Value <= 0 || ln.Element(Ns.A + "noFill") != null)
                return;
            image.FrameWidthPt = OoxmlUtil.EmuToPoints(w.Value);
            var fill = ln.Element(Ns.A + "solidFill");
            var clr = fill != null ? fill.Element(Ns.A + "srgbClr") : null;
            image.FrameColor = clr != null ? OoxmlUtil.Str(clr, "val") : null;
        }

        /// <summary>Reads the picture's source crop (a:srcRect, thousandths of a percent per edge).</summary>
        private static void ReadCrop(XElement scope, ImageInline image)
        {
            var srcRect = scope.Descendants(Ns.A + "srcRect").FirstOrDefault();
            if (srcRect == null)
                return;
            image.CropLeft = (OoxmlUtil.Dbl(srcRect, "l") ?? 0) / 100000.0;
            image.CropTop = (OoxmlUtil.Dbl(srcRect, "t") ?? 0) / 100000.0;
            image.CropRight = (OoxmlUtil.Dbl(srcRect, "r") ?? 0) / 100000.0;
            image.CropBottom = (OoxmlUtil.Dbl(srcRect, "b") ?? 0) / 100000.0;
        }

        private ImageInline CreateImage(string relId, double widthPt, double heightPt,
                                        double rotation, string description, CharacterFormat format)
        {
            if (string.IsNullOrEmpty(relId))
                return null;
            var rel = _package.GetRelationship(_currentPart, relId);
            if (rel == null || rel.External)
            {
                if (rel != null)
                    Warn("Linked (external) images are not embedded.");
                return null;
            }

            var partName = OpcPackage.ResolveTarget(_currentPart, rel.Target);
            var data = _package.ReadPart(partName);
            if (data == null)
                return null;

            return new ImageInline
            {
                PartName = partName,
                Data = data,
                WidthPt = widthPt,
                HeightPt = heightPt,
                RotationDeg = rotation,
                Description = description,
                Format = format,
                LinkUrl = _currentHyperlink,
            };
        }

        // ---------------------------------------------------------------- tables

        private Table ReadTable(XElement el)
        {
            var table = new Table();
            var tblPr = el.Element(Ns.W + "tblPr");
            var styleId = OoxmlUtil.ChildVal(tblPr, Ns.W + "tblStyle");

            // Table style contributions first, direct properties last.
            if (styleId != null)
            {
                foreach (var stylePr in _styles.TableStyleProperties(styleId))
                    ApplyTableProperties(table, stylePr);
            }
            ApplyTableProperties(table, tblPr);

            var grid = el.Element(Ns.W + "tblGrid");
            if (grid != null)
            {
                foreach (var col in grid.Elements(Ns.W + "gridCol"))
                {
                    var w = OoxmlUtil.Dbl(col, Ns.W + "w") ?? 0;
                    table.Grid.Add(OoxmlUtil.TwipsToPoints(w));
                }
            }

            Borders styleCellBorders = null;
            if (styleId != null)
            {
                foreach (var tcPr in _styles.TableStyleCellProperties(styleId))
                {
                    var b = FormatReader.ReadBorders(tcPr.Element(Ns.W + "tcBorders"));
                    if (b != null)
                        styleCellBorders = b;
                }
            }

            var scoped = false;
            if (styleId != null)
            {
                _tableStyleScope.Add(new KeyValuePair<ParagraphFormat, CharacterFormat>(
                    _styles.ParagraphFormatFor(styleId), _styles.CharacterFormatFor(styleId)));
                scoped = true;
            }

            try
            {
                foreach (var rowEl in OoxmlUtil.EffectiveElements(el))
                {
                    if (rowEl.Name != Ns.W + "tr")
                        continue;
                    table.Rows.Add(ReadRow(rowEl, table, styleCellBorders));
                }
            }
            finally
            {
                if (scoped)
                    _tableStyleScope.RemoveAt(_tableStyleScope.Count - 1);
            }

            return table;
        }

        private void ApplyTableProperties(Table table, XElement tblPr)
        {
            if (tblPr == null)
                return;

            var borders = FormatReader.ReadBorders(tblPr.Element(Ns.W + "tblBorders"));
            if (borders != null)
            {
                if (table.Borders == null)
                {
                    table.Borders = borders;
                }
                else
                {
                    if (borders.Top != null) table.Borders.Top = borders.Top;
                    if (borders.Bottom != null) table.Borders.Bottom = borders.Bottom;
                    if (borders.Left != null) table.Borders.Left = borders.Left;
                    if (borders.Right != null) table.Borders.Right = borders.Right;
                    if (borders.InsideH != null) table.Borders.InsideH = borders.InsideH;
                    if (borders.InsideV != null) table.Borders.InsideV = borders.InsideV;
                }
            }

            var shd = tblPr.Element(Ns.W + "shd");
            if (shd != null)
                table.Shading = FormatReader.ReadShadingFill(shd);

            var tblW = tblPr.Element(Ns.W + "tblW");
            if (tblW != null)
            {
                var type = (OoxmlUtil.Str(tblW, Ns.W + "type") ?? "auto").ToLowerInvariant();
                var w = OoxmlUtil.Dbl(tblW, Ns.W + "w") ?? 0;
                if (type == "pct")
                {
                    table.PreferredWidthIsPercent = true;
                    table.PreferredWidthPt = w / 50.0;   // fiftieths of a percent
                }
                else if (type == "dxa" && w > 0)
                {
                    table.PreferredWidthIsPercent = false;
                    table.PreferredWidthPt = OoxmlUtil.TwipsToPoints(w);
                }
            }

            var jc = OoxmlUtil.ChildVal(tblPr, Ns.W + "jc");
            if (jc != null)
            {
                switch (jc.ToLowerInvariant())
                {
                    case "center": table.Alignment = TextAlignment.Center; break;
                    case "right":
                    case "end": table.Alignment = TextAlignment.Right; break;
                    default: table.Alignment = TextAlignment.Left; break;
                }
            }

            var ind = tblPr.Element(Ns.W + "tblInd");
            var indW = OoxmlUtil.Dbl(ind, Ns.W + "w");
            if (indW.HasValue)
                table.IndentPt = OoxmlUtil.TwipsToPoints(indW.Value);

            var margins = tblPr.Element(Ns.W + "tblCellMar");
            if (margins != null)
            {
                var left = CellMargin(margins, "left") ?? CellMargin(margins, "start");
                if (left.HasValue) table.CellMarginLeftPt = left.Value;
                var right = CellMargin(margins, "right") ?? CellMargin(margins, "end");
                if (right.HasValue) table.CellMarginRightPt = right.Value;
                var top = CellMargin(margins, "top");
                if (top.HasValue) table.CellMarginTopPt = top.Value;
                var bottom = CellMargin(margins, "bottom");
                if (bottom.HasValue) table.CellMarginBottomPt = bottom.Value;
            }

            var spacing = tblPr.Element(Ns.W + "tblCellSpacing");
            var spacingW = OoxmlUtil.Dbl(spacing, Ns.W + "w");
            if (spacingW.HasValue)
                table.CellSpacingPt = OoxmlUtil.TwipsToPoints(spacingW.Value);
        }

        private static double? CellMargin(XElement parent, string side)
        {
            var el = parent.Element(Ns.W + side);
            if (el == null)
                return null;
            var w = OoxmlUtil.Dbl(el, Ns.W + "w");
            if (!w.HasValue)
                return null;
            var type = (OoxmlUtil.Str(el, Ns.W + "type") ?? "dxa").ToLowerInvariant();
            return type == "dxa" ? OoxmlUtil.TwipsToPoints(w.Value) : w.Value;
        }

        private TableRow ReadRow(XElement rowEl, Table table, Borders styleCellBorders)
        {
            var row = new TableRow();
            var trPr = rowEl.Element(Ns.W + "trPr");
            if (trPr != null)
            {
                row.IsHeader = OoxmlUtil.Toggle(trPr, Ns.W + "tblHeader") == true;
                row.CantSplit = OoxmlUtil.Toggle(trPr, Ns.W + "cantSplit") == true;
                var height = trPr.Element(Ns.W + "trHeight");
                var hVal = OoxmlUtil.Dbl(height, Ns.W + "val");
                if (hVal.HasValue)
                {
                    // The schema default for w:hRule is "auto", but Word in practice treats a
                    // trHeight without hRule as a minimum height ("atLeast") — verified against
                    // Word's own rendering of documents that omit the attribute.
                    var rule = (OoxmlUtil.Str(height, Ns.W + "hRule") ?? "atleast").ToLowerInvariant();
                    if (rule == "exact")
                    {
                        row.HeightPt = OoxmlUtil.TwipsToPoints(hVal.Value);
                        row.HeightExact = true;
                    }
                    else if (rule == "atleast")
                    {
                        row.HeightPt = OoxmlUtil.TwipsToPoints(hVal.Value);
                        row.HeightExact = false;
                    }
                }
            }

            foreach (var cellEl in OoxmlUtil.EffectiveElements(rowEl))
            {
                if (cellEl.Name != Ns.W + "tc")
                    continue;
                var cell = ReadCell(cellEl, table, styleCellBorders);

                // Legacy horizontal merge (w:hMerge): a continued cell folds into the cell
                // it continues, widening that cell's span — the MPOB licence uses this for
                // its full-width "NAMA LADANG" rows.
                var tcPr = cellEl.Element(Ns.W + "tcPr");
                var hMerge = tcPr == null ? null : tcPr.Element(Ns.W + "hMerge");
                if (hMerge != null)
                {
                    var val = (OoxmlUtil.Str(hMerge, Ns.W + "val") ?? "continue").ToLowerInvariant();
                    if (val != "restart" && row.Cells.Count > 0)
                    {
                        var previous = row.Cells[row.Cells.Count - 1];
                        previous.GridSpan = Math.Max(1, previous.GridSpan) + Math.Max(1, cell.GridSpan);
                        continue;
                    }
                }
                row.Cells.Add(cell);
            }
            return row;
        }

        private TableCell ReadCell(XElement cellEl, Table table, Borders styleCellBorders)
        {
            var cell = new TableCell();
            var tcPr = cellEl.Element(Ns.W + "tcPr");
            if (styleCellBorders != null)
                cell.Borders = styleCellBorders.Clone();

            if (tcPr != null)
            {
                var span = OoxmlUtil.ChildVal(tcPr, Ns.W + "gridSpan");
                int gs;
                if (span != null && int.TryParse(span, out gs) && gs > 0)
                    cell.GridSpan = gs;

                var tcW = tcPr.Element(Ns.W + "tcW");
                if (tcW != null)
                {
                    var type = (OoxmlUtil.Str(tcW, Ns.W + "type") ?? "auto").ToLowerInvariant();
                    var w = OoxmlUtil.Dbl(tcW, Ns.W + "w") ?? 0;
                    if (type == "pct")
                    {
                        cell.WidthIsPercent = true;
                        cell.WidthPt = w / 50.0;
                    }
                    else if (type == "dxa")
                    {
                        cell.WidthPt = OoxmlUtil.TwipsToPoints(w);
                    }
                }

                var vMerge = tcPr.Element(Ns.W + "vMerge");
                if (vMerge != null)
                {
                    var val = (OoxmlUtil.Str(vMerge, Ns.W + "val") ?? "continue").ToLowerInvariant();
                    cell.VMerge = val == "restart" ? VerticalMerge.Restart : VerticalMerge.Continue;
                }

                var borders = FormatReader.ReadBorders(tcPr.Element(Ns.W + "tcBorders"));
                if (borders != null)
                {
                    if (cell.Borders == null)
                    {
                        cell.Borders = borders;
                    }
                    else
                    {
                        if (borders.Top != null) cell.Borders.Top = borders.Top;
                        if (borders.Bottom != null) cell.Borders.Bottom = borders.Bottom;
                        if (borders.Left != null) cell.Borders.Left = borders.Left;
                        if (borders.Right != null) cell.Borders.Right = borders.Right;
                    }
                }

                var shd = tcPr.Element(Ns.W + "shd");
                if (shd != null)
                    cell.Shading = FormatReader.ReadShadingFill(shd);

                var vAlign = OoxmlUtil.ChildVal(tcPr, Ns.W + "vAlign");
                if (vAlign != null)
                {
                    switch (vAlign.ToLowerInvariant())
                    {
                        case "center": cell.VAlign = VerticalCellAlignment.Center; break;
                        case "bottom": cell.VAlign = VerticalCellAlignment.Bottom; break;
                        default: cell.VAlign = VerticalCellAlignment.Top; break;
                    }
                }

                var direction = OoxmlUtil.ChildVal(tcPr, Ns.W + "textDirection");
                if (direction != null && (direction == "btLr" || direction == "tbRl"))
                {
                    cell.Vertical = true;
                    Warn("Vertical text in table cells is rendered horizontally.");
                }

                var cellMar = tcPr.Element(Ns.W + "tcMar");
                if (cellMar != null)
                {
                    cell.MarginLeftPt = CellMargin(cellMar, "left") ?? CellMargin(cellMar, "start");
                    cell.MarginRightPt = CellMargin(cellMar, "right") ?? CellMargin(cellMar, "end");
                    cell.MarginTopPt = CellMargin(cellMar, "top");
                    cell.MarginBottomPt = CellMargin(cellMar, "bottom");
                }
            }

            cell.Blocks = ReadBlocks(cellEl);
            if (cell.Blocks.Count == 0)
                cell.Blocks.Add(new Paragraph { Format = _styles.DefaultParagraphFormat.Clone(), RunDefaults = _styles.DefaultCharacterFormat.Clone() });
            return cell;
        }

        // ---------------------------------------------------------------- sections

        private SectionProperties ReadSectionProperties(XElement sectPr, SectionProperties inheritFrom)
        {
            // Headers and footers are inherited from the previous section unless this one declares its own.
            var props = inheritFrom != null ? inheritFrom.Clone() : new SectionProperties();
            // Page numbering settings are per section; they are not inherited.
            props.PageNumberStart = null;
            props.PageNumberFormat = null;
            var declaresHeader = sectPr.Element(Ns.W + "headerReference") != null;
            var declaresFooter = sectPr.Element(Ns.W + "footerReference") != null;
            if (declaresHeader)
                props.HeaderDefault = props.HeaderFirst = props.HeaderEven = null;
            if (declaresFooter)
                props.FooterDefault = props.FooterFirst = props.FooterEven = null;

            var pgSz = sectPr.Element(Ns.W + "pgSz");
            if (pgSz != null)
            {
                var w = OoxmlUtil.Dbl(pgSz, Ns.W + "w");
                var h = OoxmlUtil.Dbl(pgSz, Ns.W + "h");
                if (w.HasValue && w.Value > 0) props.PageWidthPt = OoxmlUtil.TwipsToPoints(w.Value);
                if (h.HasValue && h.Value > 0) props.PageHeightPt = OoxmlUtil.TwipsToPoints(h.Value);
                props.Landscape = string.Equals(OoxmlUtil.Str(pgSz, Ns.W + "orient"), "landscape", StringComparison.OrdinalIgnoreCase);
            }

            var pgMar = sectPr.Element(Ns.W + "pgMar");
            if (pgMar != null)
            {
                var left = OoxmlUtil.Dbl(pgMar, Ns.W + "left");
                var right = OoxmlUtil.Dbl(pgMar, Ns.W + "right");
                var top = OoxmlUtil.Dbl(pgMar, Ns.W + "top");
                var bottom = OoxmlUtil.Dbl(pgMar, Ns.W + "bottom");
                var header = OoxmlUtil.Dbl(pgMar, Ns.W + "header");
                var footer = OoxmlUtil.Dbl(pgMar, Ns.W + "footer");
                var gutter = OoxmlUtil.Dbl(pgMar, Ns.W + "gutter");
                if (left.HasValue) props.MarginLeftPt = OoxmlUtil.TwipsToPoints(left.Value);
                if (right.HasValue) props.MarginRightPt = OoxmlUtil.TwipsToPoints(right.Value);
                if (top.HasValue) props.MarginTopPt = Math.Abs(OoxmlUtil.TwipsToPoints(top.Value));
                if (bottom.HasValue) props.MarginBottomPt = Math.Abs(OoxmlUtil.TwipsToPoints(bottom.Value));
                if (header.HasValue) props.HeaderDistancePt = OoxmlUtil.TwipsToPoints(header.Value);
                if (footer.HasValue) props.FooterDistancePt = OoxmlUtil.TwipsToPoints(footer.Value);
                if (gutter.HasValue) props.GutterPt = OoxmlUtil.TwipsToPoints(gutter.Value);
            }

            props.TitlePage = OoxmlUtil.Toggle(sectPr, Ns.W + "titlePg") == true;

            // Page borders belong to the section that declares them; they are not inherited.
            props.PageBorders = null;
            props.PageBordersDisplay = null;
            var pgBorders = sectPr.Element(Ns.W + "pgBorders");
            if (pgBorders != null)
            {
                props.PageBorders = FormatReader.ReadBorders(pgBorders);
                props.PageBordersFromText = string.Equals(OoxmlUtil.Str(pgBorders, Ns.W + "offsetFrom"), "text",
                                                          StringComparison.OrdinalIgnoreCase);
                props.PageBordersDisplay = OoxmlUtil.Str(pgBorders, Ns.W + "display");
            }

            var cols = sectPr.Element(Ns.W + "cols");
            if (cols != null)
            {
                var num = OoxmlUtil.Int(cols, Ns.W + "num");
                if (num.HasValue && num.Value > 1)
                {
                    props.Columns = num.Value;
                    var space = OoxmlUtil.Dbl(cols, Ns.W + "space");
                    if (space.HasValue)
                        props.ColumnSpacingPt = OoxmlUtil.TwipsToPoints(space.Value);
                }
                else
                {
                    props.Columns = 1;
                }
            }

            var pgNumType = sectPr.Element(Ns.W + "pgNumType");
            var start = OoxmlUtil.Int(pgNumType, Ns.W + "start");
            if (start.HasValue)
                props.PageNumberStart = start.Value;
            var numberFormat = OoxmlUtil.Str(pgNumType, Ns.W + "fmt");
            if (!string.IsNullOrEmpty(numberFormat))
                props.PageNumberFormat = numberFormat;

            foreach (var refEl in sectPr.Elements())
            {
                var isHeader = refEl.Name == Ns.W + "headerReference";
                var isFooter = refEl.Name == Ns.W + "footerReference";
                if (!isHeader && !isFooter)
                    continue;

                var relId = OoxmlUtil.Str(refEl, Ns.R + "id");
                var type = (OoxmlUtil.Str(refEl, Ns.W + "type") ?? "default").ToLowerInvariant();
                var content = ReadHeaderFooterPart(relId);
                if (content == null)
                    continue;

                if (isHeader)
                {
                    if (type == "first") props.HeaderFirst = content;
                    else if (type == "even") props.HeaderEven = content;
                    else props.HeaderDefault = content;
                }
                else
                {
                    if (type == "first") props.FooterFirst = content;
                    else if (type == "even") props.FooterEven = content;
                    else props.FooterDefault = content;
                }
            }

            return props;
        }

        private HeaderFooter ReadHeaderFooterPart(string relId)
        {
            if (string.IsNullOrEmpty(relId))
                return null;
            var rel = _package.GetRelationship(_documentPart, relId);
            if (rel == null || rel.External)
                return null;
            var partName = OpcPackage.ResolveTarget(_documentPart, rel.Target);
            var doc = _package.ReadXml(partName);
            if (doc == null || doc.Root == null)
                return null;

            var previousPart = _currentPart;
            _currentPart = partName;
            try
            {
                var hf = new HeaderFooter { Blocks = ReadBlocks(doc.Root) };
                return hf.Blocks.Count == 0 ? null : hf;
            }
            finally
            {
                _currentPart = previousPart;
            }
        }
    }
}
