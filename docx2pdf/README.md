# Docx2Pdf

A dependency-free DOCX → PDF converter written in C# for **.NET Standard 2.0**.

No Microsoft Word, no COM interop, no native libraries, no NuGet packages — the library reads the
OOXML package, lays the document out itself and writes the PDF byte for byte.

## Which runtimes can use it

The library targets `netstandard2.0`, which is consumable from:

| Platform | Supported versions |
| --- | --- |
| .NET Framework | 4.6.1 and later (4.6.1, 4.6.2, 4.7.x, 4.8, 4.8.1) |
| .NET Core | 2.0 and later |
| .NET | 5, 6, 7, 8, 9, 10 (and later) |
| Mono / Xamarin / Unity | Mono 5.4+, Xamarin.iOS 10.14+, Xamarin.Android 8.0+, Unity 2018.1+ |

Works on Windows, Linux and macOS. On .NET Framework 4.6.0 and earlier, netstandard2.0 assemblies
cannot be referenced — build the sources directly into your project instead (the code uses no APIs
newer than .NET Framework 4.5 except `System.IO.Compression`).

## Usage

```csharp
using Docx2Pdf;

// Simplest form
DocxToPdfConverter.Convert(@"C:\docs\manual.docx", @"C:\docs\manual.pdf");

// With options and a result report
var options = new ConversionOptions
{
    EmbedFonts = true,          // embed the document's real fonts (default)
    GenerateOutline = true,     // PDF bookmarks from Heading 1-9
    CreateHyperlinks = true,    // clickable links
    DefaultFontFamily = "Calibri",
};

ConversionResult result = DocxToPdfConverter.Convert(inputPath, outputPath, options);
Console.WriteLine($"{result.PageCount} pages, {result.ByteCount:N0} bytes");
foreach (var warning in result.Warnings)
    Console.WriteLine("warning: " + warning);
```

Streams and byte arrays work too:

```csharp
using (var input = File.OpenRead("in.docx"))
using (var output = File.Create("out.pdf"))
    DocxToPdfConverter.Convert(input, output);

byte[] pdfBytes = DocxToPdfConverter.Convert(File.ReadAllBytes("in.docx"));
```

`Convert` is thread-safe in the sense that separate calls do not share mutable state; the system
font index is built once per process and guarded by a lock.

### Command line

```
docx2pdf <input.docx> [output.pdf] [options]

  --no-embed           substitute standard PDF fonts instead of embedding font files
  --no-images          skip images
  --no-compress        write uncompressed streams (useful for inspecting the output)
  --no-outline         do not generate PDF bookmarks from headings
  --no-links           do not generate link annotations
  --no-headers         skip page headers and footers
  --font-dir=PATH      additional directory to search for font files
  --default-font=NAME  font used when the document does not name one
  --max-pages=N        stop after N pages
  --quiet              print nothing on success
```

## What is supported

**Text and paragraphs** — styles with `basedOn` inheritance and document defaults, bold/italic,
underline (single, double, dotted, dashed, thick), strike-through, super/subscript, small caps,
all caps, hidden text, character spacing, text colour, highlighting, run and paragraph shading,
paragraph borders, alignment (left/centre/right/justified), first-line and hanging indents,
negative indents, space before/after, contextual spacing, line spacing (multiple, at-least, exact),
tab stops (left/centre/right/decimal with dot, hyphen and underscore leaders), line and page breaks.

**Lists** — `numbering.xml` with abstract definitions, level overrides, `numStyleLink`,
multi-level labels (`%1.%2.`), decimal, letters, roman, ordinal, Chinese and bullet formats,
level-to-style binding through `w:lvl/w:pStyle` (which is how heading numbering such as
*4.1, 4.2* is defined), per-level indents and bullet fonts.

**Tables** — table styles (including the `w:pPr`/`w:rPr` they impose on every paragraph in the
table, which is what makes cell text single spaced with no space after),
grids and column widths (dxa/pct/auto), `gridSpan`, cell margins,
cell shading, table and cell borders with inside/outside resolution, vertical alignment,
header rows repeated on every page, and **rows split across pages** like Word.
Row heights follow `w:trHeight/@w:hRule` — `exact`, `atLeast`, and `auto`. When the attribute is
absent the schema default is `auto`, but Word treats the bare value as a minimum height
(`atLeast`); this library follows Word, verified by page-image comparison against Word's output.
Vertically merged cells are measured across every row they span, growing the last row when needed.

**Page setup** — multiple sections, page size and orientation, margins, page borders,
headers and footers (default/first/even, inherited between sections), header-aware body area
(the body moves down when the header is taller than the top margin), `PAGE`/`NUMPAGES` fields
with per-section number formats (decimal, roman, letters), widow/orphan control, keep-with-next
and keep-lines-together.

**Graphics** — DrawingML and legacy VML images, PNG (all colour types, bit depths and Adam7
interlacing, with alpha as a soft mask), JPEG (embedded without re-encoding), BMP, GIF, image
rotation, and text boxes. Floating (anchored) pictures and text boxes are positioned the way the
document asks: `positionH`/`positionV` offsets and alignments relative to the page, the margins,
the column or the anchoring paragraph, including `behindDoc` draw order.

**Fonts** — the document's own fonts are located on the machine and embedded as CID (Identity-H)
composite fonts with a `ToUnicode` CMap, so text is selectable and searchable. TrueType, OpenType
(CFF) and TrueType collections (`.ttc`) are supported. Characters the requested font cannot render
fall back to other installed fonts, so CJK text in a Latin-font document still renders. When no
font file can be found, the base-14 PDF fonts are substituted with correct Adobe metrics.

**PDF output** — Flate compression, document outline (bookmarks) from headings, internal and
external link annotations, bookmarks as link targets, document metadata.

## Known limitations

* Floating objects are positioned correctly but text does not wrap around them: `square`/`tight`
  wrapping is treated as `wrapNone`, so a float overlays (or sits behind) the text instead of
  pushing it aside.
* Vector drawings, WordArt, charts and SmartArt are not rendered (a warning is reported).
  EMF/WMF/TIFF/WDP images are skipped for the same reason.
* Footnotes and endnotes are collected at the end of the document instead of the foot of the page.
* Multi-column sections are laid out as a single column.
* Right-to-left text is rendered left-to-right; complex-script shaping (Arabic, Indic) is not done.
* Embedded fonts are not subsetted, so documents using many fonts produce larger files.
* Pagination is very close to Word but not always identical; small differences in line metrics can
  shift a page break by a line.

## Project layout

```
docx2pdf/
  src/Docx2Pdf/            the library (netstandard2.0, no dependencies)
    Ooxml/                 OPC package reader, styles, numbering, document reader
    Model/                 the document model produced by the reader
    Layout/                line breaking, tables, pagination, headers/footers
    Fonts/                 TrueType parsing, system font index, PDF font objects
    Images/                PNG/JPEG/BMP/GIF decoding
    Pdf/                   PDF object model, writer, content-stream renderer
  samples/Docx2Pdf.Cli/    command line front end
  tests/Docx2Pdf.Tests/    self-contained test runner (no test framework needed)
  tools/Docx2Pdf.Compare/  development tool: rasterises PDFs and diffs pages (uses PDFium)
```

Only `src/Docx2Pdf` ships. The tools and tests are development aids; the comparison tool is the
only project that uses a NuGet package.

## Building

```powershell
dotnet build docx2pdf\Docx2Pdf.sln -c Release
```

## Testing

`.sample\test-docx2pdf.ps1` builds everything, runs the test suite, converts every `.docx`
in `.sample`, and renders each PDF page to PNG under `.sample\_render`:

```powershell
.\.sample\test-docx2pdf.ps1
```

With `-Compare` it additionally exports a reference PDF from Microsoft Word and compares the two
documents page by page, writing `word` / `ours` / `diff` images and a similarity percentage into
`.sample\_compare`:

```powershell
.\.sample\test-docx2pdf.ps1 -Compare -Dpi 150
```

The diff image is red where Word puts ink and we do not, and blue where we put ink and Word does
not. Word automation needs an interactive desktop session; if Word does not respond within the
timeout the script warns and skips the comparison.

When Word automation is unavailable, screenshots of the Word pages work just as well. Save them
next to the script as `01.png`, `02.png`, ... and run:

```powershell
.\.sample\test-docx2pdf.ps1 -Screenshots showlicenseindocx
```

Each screenshot is aligned with our page on the page border rectangle, then every text line is
matched by position and its vertical difference reported, ending with the mean and largest error
in points. That number is the regression metric to watch when changing layout code.
