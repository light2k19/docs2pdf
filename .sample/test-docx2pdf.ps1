<#
.SYNOPSIS
    Builds the Docx2Pdf library, runs its tests, converts every .docx in this folder to PDF,
    renders each PDF page to PNG and - when Microsoft Word is available - compares our output
    with Word's own PDF export page by page.

.DESCRIPTION
    Outputs produced next to this script:
      <name>.pdf              our conversion
      _render\*.png           page images of our PDF
      _reference\*.word.pdf   Word's export of the same document   (-Compare only)
      _compare\<name>\*.png   word / ours / diff page images        (-Compare only)

    The diff images are red where Word has ink we do not, and blue where we draw ink Word does not.

.PARAMETER Compare
    Also export reference PDFs with Microsoft Word and produce a page-by-page comparison.
    Requires Word to be installed and an interactive desktop session.

.PARAMETER Screenshots
    Base name of a document for which page screenshots taken from Word (01.png, 02.png, ...) sit
    next to this script. Each screenshot is aligned with the matching page of our PDF on the page
    border, and every text line is compared; the report lists each line's vertical difference and
    a red/blue overlay is written to _compare.

.PARAMETER Dpi
    Rasterisation resolution for the page images (default 96).

.EXAMPLE
    .\test-docx2pdf.ps1
    .\test-docx2pdf.ps1 -Compare -Dpi 150
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$SkipTests,
    [switch]$Compare,
    [string]$Screenshots,
    [int]$Dpi = 96,
    [int]$WordTimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'
$sampleDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$repoRoot = Join-Path (Split-Path $sampleDir -Parent) 'docx2pdf'

$solution = Join-Path $repoRoot 'Docx2Pdf.sln'
$cli = Join-Path $repoRoot 'samples\Docx2Pdf.Cli\bin\Release\net10.0\Docx2Pdf.Cli.exe'
$tests = Join-Path $repoRoot 'tests\Docx2Pdf.Tests\bin\Release\net10.0\Docx2Pdf.Tests.exe'
$compareTool = Join-Path $repoRoot 'tools\Docx2Pdf.Compare\bin\Release\net10.0\Docx2Pdf.Compare.exe'

$renderDir = Join-Path $sampleDir '_render'
$referenceDir = Join-Path $sampleDir '_reference'
$compareDir = Join-Path $sampleDir '_compare'

function Write-Section($text) {
    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkGray
    Write-Host $text -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor DarkGray
}

# ---------------------------------------------------------------- build

if (-not $SkipBuild) {
    Write-Section 'Building Docx2Pdf'
    dotnet build $solution -c Release --nologo -v minimal
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }
}

foreach ($required in @($cli, $tests)) {
    if (-not (Test-Path $required)) { throw "Missing build output: $required (run without -SkipBuild)" }
}

# ---------------------------------------------------------------- unit tests

if (-not $SkipTests) {
    Write-Section 'Running tests'
    & $tests $sampleDir $sampleDir
    if ($LASTEXITCODE -ne 0) { throw "Tests failed with exit code $LASTEXITCODE" }
}

# ---------------------------------------------------------------- conversion

Write-Section 'Converting documents'
$documents = Get-ChildItem -Path $sampleDir -Filter *.docx | Where-Object { -not $_.Name.StartsWith('~$') }
if ($documents.Count -eq 0) { throw "No .docx files found in $sampleDir" }

$results = @()
foreach ($doc in $documents) {
    $pdf = Join-Path $sampleDir ($doc.BaseName + '.pdf')
    $started = Get-Date
    & $cli $doc.FullName $pdf
    if ($LASTEXITCODE -ne 0) { throw "Conversion failed for $($doc.Name)" }
    $results += [pscustomobject]@{
        Document = $doc.Name
        Pdf      = $pdf
        Seconds  = [math]::Round(((Get-Date) - $started).TotalSeconds, 2)
    }
}

# ---------------------------------------------------------------- page images

if (Test-Path $compareTool) {
    Write-Section 'Rendering page images'
    New-Item -ItemType Directory -Force -Path $renderDir | Out-Null
    foreach ($result in $results) {
        & $compareTool --render $result.Pdf $renderDir $Dpi
    }
    Write-Host "Page images: $renderDir"
} else {
    Write-Warning "Compare tool not built ($compareTool); skipping page images."
}

# ---------------------------------------------------------------- screenshot comparison

if ($Screenshots) {
    Write-Section "Comparing with Word screenshots ($Screenshots)"
    $target = $results | Where-Object { [IO.Path]::GetFileNameWithoutExtension($_.Pdf) -eq $Screenshots } | Select-Object -First 1
    if (-not $target) {
        Write-Warning "No converted document called '$Screenshots'."
    } else {
        New-Item -ItemType Directory -Force -Path $compareDir | Out-Null
        $shots = Get-ChildItem -Path $sampleDir -Filter '*.png' |
                 Where-Object { $_.BaseName -match '^\d+$' } | Sort-Object BaseName
        if ($shots.Count -eq 0) {
            Write-Warning "No page screenshots (01.png, 02.png, ...) found in $sampleDir."
        }
        foreach ($shot in $shots) {
            $page = [int]$shot.BaseName
            $ours = Join-Path $renderDir ("{0}-page-{1:D3}.png" -f $Screenshots, $page)
            if (-not (Test-Path $ours)) {
                Write-Warning "No rendered page $page for $Screenshots."
                continue
            }
            Write-Host ''
            Write-Host ("page {0}" -f $page) -ForegroundColor Yellow
            & $compareTool --align $shot.FullName $ours (Join-Path $compareDir ("page-{0:D3}-diff.png" -f $page))
        }
    }
}

# ---------------------------------------------------------------- Word comparison

if ($Compare) {
    Write-Section 'Exporting reference PDFs with Microsoft Word'
    New-Item -ItemType Directory -Force -Path $referenceDir | Out-Null

    $exportScript = {
        param($files, $outputDirectory)
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try {
            foreach ($file in $files) {
                $document = $word.Documents.Open($file, $false, $true, $false)
                $target = Join-Path $outputDirectory ([IO.Path]::GetFileNameWithoutExtension($file) + '.word.pdf')
                # 17 = wdExportFormatPDF
                $document.ExportAsFixedFormat($target, 17)
                $document.Close($false)
            }
        } finally {
            $word.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
    }

    $job = Start-Job -ScriptBlock $exportScript -ArgumentList (,($documents.FullName)), $referenceDir
    if (Wait-Job $job -Timeout $WordTimeoutSeconds) {
        Receive-Job $job
    } else {
        Stop-Job $job
        Write-Warning "Word did not finish within $WordTimeoutSeconds seconds. Automation is often blocked in non-interactive sessions; run this script from a normal desktop PowerShell window."
    }
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    Write-Section 'Comparing pages'
    foreach ($result in $results) {
        $reference = Join-Path $referenceDir ([IO.Path]::GetFileNameWithoutExtension($result.Pdf) + '.word.pdf')
        if (-not (Test-Path $reference)) {
            Write-Warning "No Word reference for $($result.Document); skipped."
            continue
        }
        $target = Join-Path $compareDir ([IO.Path]::GetFileNameWithoutExtension($result.Pdf))
        Write-Host ''
        Write-Host $result.Document -ForegroundColor Yellow
        & $compareTool $reference $result.Pdf $target $Dpi
    }
    Write-Host ''
    Write-Host "Comparison images: $compareDir"
}

# ---------------------------------------------------------------- summary

Write-Section 'Summary'
$results | Format-Table -AutoSize
Write-Host 'Done.' -ForegroundColor Green
