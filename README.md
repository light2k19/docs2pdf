# ai_docs

Document processing libraries in C#.

## Projects

| Project | Description |
| --- | --- |
| [docx2pdf](docx2pdf/) | A dependency-free DOCX → PDF converter for .NET Standard 2.0 — no Word, no COM interop, no native libraries, no NuGet packages. See its [README](docx2pdf/README.md). |

## Quick start

```powershell
dotnet build docx2pdf\Docx2Pdf.sln -c Release
```

```csharp
using Docx2Pdf;

DocxToPdfConverter.Convert(@"C:\docs\manual.docx", @"C:\docs\manual.pdf");
```

## Repository layout

```
docx2pdf/          the converter (library, CLI, tests, comparison tool)
.sample/           local test documents and render output — not tracked
  test-docx2pdf.ps1  build + convert + render + compare against Word
```

Everything under `.sample/` is gitignored apart from `test-docx2pdf.ps1`; put your own `.docx`
files there and run the script to exercise the converter. See
[Testing](docx2pdf/README.md#testing) for what the comparison modes do.
