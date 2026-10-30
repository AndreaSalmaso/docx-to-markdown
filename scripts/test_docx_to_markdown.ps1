$ErrorActionPreference = 'Stop'

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$converter = Join-Path $scriptDirectory 'docx_to_markdown.ps1'
$pandoc = if ($env:LOCALAPPDATA) {
    Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
}
$pandocCommand = Get-Command pandoc -ErrorAction SilentlyContinue
if ($pandocCommand) {
    $pandoc = $pandocCommand.Source
}
if (-not $pandoc -or -not (Test-Path -LiteralPath $pandoc -PathType Leaf)) {
    throw 'Pandoc non trovato: installalo oppure rendilo disponibile nel PATH per eseguire i test.'
}

$testRoot = Join-Path $env:TEMP ('docx-powershell-tests-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($testRoot)

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Write-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$Name,
        [string]$Content
    )
    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new(
            $stream,
            [System.Text.UTF8Encoding]::new($false)
        )
        try {
            $writer.Write($Content)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function New-TestDocx {
    param([string]$Path)

    $parent = Split-Path -Parent $Path
    [void][System.IO.Directory]::CreateDirectory($parent)
    $fileStream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::ReadWrite
    )
    try {
        $archive = [System.IO.Compression.ZipArchive]::new(
            $fileStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false
        )
        try {
            Write-ZipTextEntry $archive '[Content_Types].xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'@
            Write-ZipTextEntry $archive '_rels/.rels' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
            Write-ZipTextEntry $archive 'word/document.xml' @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Titolo fixture</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Sezione convertita</w:t></w:r></w:p>
    <w:p><w:r><w:t>Parole che non centrano con il contenuto. REFUSO.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>
'@
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

function Invoke-Converter {
    param([string[]]$Arguments)

    $previousErrorPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $converter @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorPreference
    }
    return [PSCustomObject]@{
        ExitCode = $exitCode
        Output = ($output | Out-String)
    }
}

try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $inputRoot = Join-Path $testRoot 'documenti con spazi'
    $source = Join-Path $inputRoot 'fixture.docx'
    New-TestDocx $source

    $correctionsFile = Join-Path $testRoot 'correzioni.json'
    [System.IO.File]::WriteAllText(
        $correctionsFile,
        '{"REFUSO":"correzione"}',
        [System.Text.UTF8Encoding]::new($false)
    )

    $run = Invoke-Converter @(
        $source,
        '-PandocPath',
        $pandoc,
        '-Corrections',
        $correctionsFile,
        '-NoExtractMedia'
    )
    Assert-True ($run.ExitCode -eq 0) "Conversione file singolo fallita: $($run.Output)"
    $markdownPath = Join-Path $inputRoot 'fixture.md'
    Assert-True (Test-Path -LiteralPath $markdownPath) 'The Markdown file was not created.'
    $markdown = [System.IO.File]::ReadAllText($markdownPath)
    Assert-True ($markdown.Contains('# Titolo fixture')) "DOCX heading was not converted to GFM: `n$markdown"
    Assert-True ($markdown.Contains('## Sezione convertita')) 'Standalone bold text was not promoted to a GFM heading.'
    Assert-True ($markdown.Contains('non c' + [char]0x2019 + 'entrano con')) 'Targeted typo cleanup was not applied.'
    Assert-True ($markdown.Contains('correzione.')) 'The JSON correction was not applied.'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $inputRoot 'fixture_media'))) 'Media folder was created with -NoExtractMedia.'
    $gfmRoundTrip = Join-Path $testRoot 'round-trip.md'
    & $pandoc --from=gfm --to=gfm -o $gfmRoundTrip $markdownPath
    Assert-True ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $gfmRoundTrip)) 'Pandoc could not parse the generated Markdown as GFM.'

    $recursiveRoot = Join-Path $testRoot 'input-ricorsivo'
    $recursiveRootDocx = Join-Path $recursiveRoot 'principale.docx'
    New-TestDocx $recursiveRootDocx
    $recursiveInput = Join-Path $recursiveRoot 'sottocartella'
    $recursiveDocx = Join-Path $recursiveInput 'secondo.docx'
    New-TestDocx $recursiveDocx
    $outputDirectory = Join-Path $testRoot 'output'
    $run = Invoke-Converter @(
        $recursiveRoot,
        '-PandocPath',
        $pandoc,
        '-Recursive',
        '-NoExtractMedia',
        '-OutputDir',
        $outputDirectory
    )
    Assert-True ($run.ExitCode -eq 0) "Conversione ricorsiva fallita: $($run.Output)"
    Assert-True (Test-Path -LiteralPath (Join-Path $outputDirectory 'principale.md')) 'Output del DOCX principale non trovato.'
    Assert-True (Test-Path -LiteralPath (Join-Path $outputDirectory 'sottocartella\secondo.md')) 'Output della sottocartella non trovato.'

    $originalOutput = [System.IO.File]::ReadAllText($markdownPath)
    $run = Invoke-Converter @($source, '-PandocPath', $pandoc, '-NoExtractMedia')
    Assert-True ($run.ExitCode -ne 0) 'An existing Markdown file was overwritten without -Force.'
    Assert-True (
        [System.IO.File]::ReadAllText($markdownPath) -ceq $originalOutput
    ) 'The Markdown changed after overwrite was refused.'

    $run = Invoke-Converter @($source, '-PandocPath', $pandoc, '-NoExtractMedia', '-Force')
    Assert-True ($run.ExitCode -eq 0) "Explicit overwrite with -Force failed: $($run.Output)"

    $skillRoot = Split-Path -Parent $scriptDirectory
    $installDestination = Join-Path $testRoot 'installed-skill'
    $staleScripts = Join-Path $installDestination 'scripts'
    [void][System.IO.Directory]::CreateDirectory($staleScripts)
    [System.IO.File]::WriteAllText((Join-Path $staleScripts 'docx_to_markdown.py'), 'obsolete')
    [System.IO.File]::WriteAllText((Join-Path $staleScripts 'test_docx_to_markdown.py'), 'obsolete')
    $installerPandocDirectory = Split-Path -Parent $pandoc
    $previousPath = $env:PATH
    try {
        $env:PATH = $installerPandocDirectory + ';' + $previousPath
        & (Join-Path $skillRoot 'install.ps1') -Destination $installDestination -Confirm:$false
        Assert-True ($LASTEXITCODE -eq 0) 'The one-command installer failed in its temporary test destination.'
    }
    finally {
        $env:PATH = $previousPath
    }
    Assert-True (Test-Path -LiteralPath (Join-Path $installDestination 'SKILL.md')) 'Installer did not copy the skill definition.'
    Assert-True (Test-Path -LiteralPath (Join-Path $installDestination 'README.md')) 'Installer did not copy the README.'
    Assert-True (Test-Path -LiteralPath (Join-Path $installDestination 'scripts\docx_to_markdown.ps1')) 'Installer did not copy the PowerShell converter.'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $installDestination 'scripts\docx_to_markdown.py'))) 'Installer left behind the obsolete Python converter.'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $installDestination 'scripts\test_docx_to_markdown.py'))) 'Installer left behind the obsolete Python tests.'

    Write-Output 'OK: DOCX-to-GFM, headings, cleanup, JSON corrections, paths with spaces, recursive mode, GFM round-trip, overwrite protection, and one-command installer.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
