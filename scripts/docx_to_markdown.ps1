<#
.SYNOPSIS
Converts DOCX documents to cleaned GitHub Flavored Markdown using Pandoc.

.DESCRIPTION
Accepts DOCX files or folders, optionally searches subfolders, extracts media,
applies deterministic Markdown cleanup, and refuses to overwrite output files
unless -Force is supplied.

.PARAMETER Inputs
One or more DOCX file paths or folders containing DOCX files.

.PARAMETER OutputDir
Optional destination folder for generated Markdown files.

.PARAMETER Recursive
Searches input folders and their subfolders for DOCX files.

.PARAMETER PandocPath
Path to the Pandoc executable. By default searches PATH and standard Windows
per-user installation locations.

.PARAMETER Corrections
Optional JSON object mapping literal text to replacement text.

.PARAMETER NoExtractMedia
Disables extraction of document images and other media.

.PARAMETER Force
Allows existing Markdown and media outputs to be replaced.

.EXAMPLE
powershell.exe -File .\docx_to_markdown.ps1 .\document.docx

.EXAMPLE
powershell.exe -File .\docx_to_markdown.ps1 .\documents -Recursive -OutputDir .\markdown
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Inputs,

    [switch]$Help,
    [string]$OutputDir,
    [switch]$Recursive,
    [string]$PandocPath,
    [string]$Corrections,
    [switch]$NoExtractMedia,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
if ($Help) {
    Get-Help $PSCommandPath -Full | Out-String | Write-Output
    exit 0
}
if (-not $Inputs -or $Inputs.Count -eq 0) {
    throw 'Specifica almeno un file DOCX o una cartella. Usa -Help per la guida.'
}

function Write-Usage {
    @'
Uso:
  .\docx_to_markdown.ps1 <file.docx|cartella> [altri percorsi] [-OutputDir cartella]
                         [-Recursive] [-PandocPath percorso] [-Corrections file.json]
                         [-NoExtractMedia] [-Force]

Converte DOCX in GitHub Flavored Markdown con Pandoc.
'@ | Write-Output
}

function Resolve-Pandoc {
    param([string]$ConfiguredPath)

    if ($ConfiguredPath) {
        if (Test-Path -LiteralPath $ConfiguredPath -PathType Leaf) {
            return (Resolve-Path -LiteralPath $ConfiguredPath).Path
        }
        $configuredCommand = Get-Command $ConfiguredPath -ErrorAction SilentlyContinue
        if ($configuredCommand) {
            return $configuredCommand.Source
        }
        throw "Pandoc non trovato nel percorso specificato: $ConfiguredPath"
    }

    $pandocCommand = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($pandocCommand) {
        return $pandocCommand.Source
    }

    if ($env:LOCALAPPDATA) {
        $userPandoc = Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
        if (Test-Path -LiteralPath $userPandoc -PathType Leaf) {
            return $userPandoc
        }
    }

    throw 'Pandoc non e nel PATH. Installalo con winget oppure specifica -PandocPath.'
}

function Read-Corrections {
    param([string]$Path)

    $correctionsList = [System.Collections.Generic.List[object]]::new()
    if (-not $Path) {
        return ,$correctionsList
    }

    try {
        $json = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path) |
            ConvertFrom-Json
    }
    catch {
        throw "Impossibile leggere il file JSON delle correzioni '$Path': $($_.Exception.Message)"
    }

    if ($json -isnot [System.Management.Automation.PSCustomObject]) {
        throw 'Il file delle correzioni deve contenere un oggetto JSON.'
    }

    foreach ($property in $json.PSObject.Properties) {
        if (
            [string]::IsNullOrEmpty($property.Name) -or
            $property.Value -isnot [string]
        ) {
            throw 'Le correzioni JSON devono associare chiavi non vuote a valori stringa.'
        }
        $correctionsList.Add(
            [PSCustomObject]@{
                Original = $property.Name
                Replacement = $property.Value
            }
        )
    }

    return ,$correctionsList
}

function Clean-Prose {
    param(
        [string]$Text,
        [System.Collections.Generic.List[object]]$CorrectionsList
    )

    $cleaned = [regex]::Replace($Text, '</?span\b[^>]*>', '', 'IgnoreCase')
    $cleaned = [regex]::Replace($cleaned, '</?u\b[^>]*>', '', 'IgnoreCase')
    $cleaned = [regex]::Replace($cleaned, '<br\s*/?>', "  `n", 'IgnoreCase')
    $cleaned = $cleaned.Replace([string][char]0x00AD, '').Replace([string][char]0x200B, '').Replace([string][char]0xFEFF, '')
    $cleaned = [regex]::Replace(
        $cleaned,
        '\b(non\s+)centrano(\s+con\b)',
        [System.Text.RegularExpressions.MatchEvaluator]{
            param($match)
            return $match.Groups[1].Value + 'c' + [char]0x2019 + 'entrano' + $match.Groups[2].Value
        },
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    foreach ($correction in $CorrectionsList) {
        $cleaned = $cleaned.Replace($correction.Original, $correction.Replacement)
    }
    return $cleaned
}

function Normalize-BoldHeadings {
    param([string]$Markdown)

    $hadTrailingNewline = $Markdown.EndsWith("`n") -or $Markdown.EndsWith("`r")
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($line in ($Markdown -split "`n")) {
        $lines.Add($line.TrimEnd("`r"))
    }
    if ($hadTrailingNewline -and $lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') {
        $lines.RemoveAt($lines.Count - 1)
    }

    $firstContentLine = -1
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if (-not [string]::IsNullOrWhiteSpace($lines[$index])) {
            $firstContentLine = $index
            break
        }
    }

    $insideFence = $false
    $fenceCharacter = [char]0
    $fenceLength = 0
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $line = $lines[$index]
        $fence = [regex]::Match($line, '^ {0,3}(`{3,}|~{3,})')
        if ($fence.Success) {
            $marker = $fence.Groups[1].Value
            if (-not $insideFence) {
                $insideFence = $true
                $fenceCharacter = $marker[0]
                $fenceLength = $marker.Length
            }
            elseif ($marker[0] -eq $fenceCharacter -and $marker.Length -ge $fenceLength) {
                $insideFence = $false
                $fenceCharacter = [char]0
                $fenceLength = 0
            }
            continue
        }
        if ($insideFence) {
            continue
        }

        $boldHeading = [regex]::Match(
            $line,
            '^(?<indent> {0,3})\*\*(?<title>[^*\n]+)\*\*[ \t]*$'
        )
        if (-not $boldHeading.Success) {
            continue
        }

        $title = $boldHeading.Groups['title'].Value.Trim()
        if (
            [string]::IsNullOrEmpty($title) -or
            $title.Length -gt 90 -or
            $title -match '[.!;:]$'
        ) {
            continue
        }
        $level = if ($index -eq $firstContentLine) { '#' } else { '##' }
        $lines[$index] = $boldHeading.Groups['indent'].Value + $level + ' ' + $title
    }

    $result = $lines -join "`n"
    if ($hadTrailingNewline) {
        $result += "`n"
    }
    return $result
}

function Clean-Markdown {
    param(
        [string]$Markdown,
        [System.Collections.Generic.List[object]]$CorrectionsList
    )

    $hadTrailingNewline = $Markdown.EndsWith("`n") -or $Markdown.EndsWith("`r")
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($line in ($Markdown -split "`n")) {
        $lines.Add($line.TrimEnd("`r"))
    }
    if ($hadTrailingNewline -and $lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') {
        $lines.RemoveAt($lines.Count - 1)
    }

    $output = [System.Collections.Generic.List[string]]::new()
    $prose = [System.Collections.Generic.List[string]]::new()
    $insideFence = $false
    $fenceCharacter = [char]0
    $fenceLength = 0

    foreach ($line in $lines) {
        $fence = [regex]::Match($line, '^ {0,3}(`{3,}|~{3,})')
        if ($fence.Success) {
            $marker = $fence.Groups[1].Value
            if (-not $insideFence) {
                if ($prose.Count -gt 0) {
                    $output.Add((Clean-Prose ($prose -join "`n") $CorrectionsList))
                    $prose.Clear()
                }
                $insideFence = $true
                $fenceCharacter = $marker[0]
                $fenceLength = $marker.Length
                $output.Add($line)
                continue
            }
            if ($marker[0] -eq $fenceCharacter -and $marker.Length -ge $fenceLength) {
                $output.Add($line)
                $insideFence = $false
                $fenceCharacter = [char]0
                $fenceLength = 0
                continue
            }
        }

        if ($insideFence) {
            $output.Add($line)
        }
        else {
            $prose.Add($line)
        }
    }

    if ($prose.Count -gt 0) {
        $output.Add((Clean-Prose ($prose -join "`n") $CorrectionsList))
    }

    $cleaned = $output -join "`n"
    if ($hadTrailingNewline) {
        $cleaned += "`n"
    }
    return Normalize-BoldHeadings $cleaned
}

function ConvertTo-ProcessArgument {
    param([string]$Value)

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }
        if ($character -eq '"') {
            [void]$builder.Append(('\' * (2 * $backslashCount + 1)))
            [void]$builder.Append('"')
            $backslashCount = 0
            continue
        }
        if ($backslashCount -gt 0) {
            [void]$builder.Append(('\' * $backslashCount))
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }
    if ($backslashCount -gt 0) {
        [void]$builder.Append(('\' * (2 * $backslashCount)))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Invoke-Pandoc {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $Executable
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $startInfo.StandardErrorEncoding = [System.Text.UTF8Encoding]::new($false)
    $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join ' ')

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw 'Windows non e riuscito ad avviare Pandoc.'
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        return [PSCustomObject]@{
            ExitCode = $process.ExitCode
            StdOut = $stdoutTask.Result
            StdErr = $stderrTask.Result
        }
    }
    catch {
        throw "Impossibile eseguire Pandoc '$Executable': $($_.Exception.Message)"
    }
    finally {
        $process.Dispose()
    }
}

function Get-ConversionJobs {
    param(
        [string[]]$InputPaths,
        [string]$DestinationDirectory,
        [bool]$SearchRecursively
    )

    $jobs = [System.Collections.Generic.List[object]]::new()
    foreach ($inputPath in $InputPaths) {
        if (-not (Test-Path -LiteralPath $inputPath)) {
            throw "Percorso inesistente: $inputPath"
        }
        $item = Get-Item -LiteralPath $inputPath
        if ($item.PSIsContainer) {
            if ($SearchRecursively) {
                $documents = Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.docx' -Recurse
            }
            else {
                $documents = Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.docx'
            }
            foreach ($document in ($documents | Sort-Object FullName)) {
                if ($DestinationDirectory) {
                    $relativePath = $document.FullName.Substring($item.FullName.TrimEnd('\').Length).TrimStart('\')
                    $destination = Join-Path $DestinationDirectory ([System.IO.Path]::ChangeExtension($relativePath, '.md'))
                }
                else {
                    $destination = [System.IO.Path]::ChangeExtension($document.FullName, '.md')
                }
                $jobs.Add(
                    [PSCustomObject]@{
                        Source = $document.FullName
                        Destination = [System.IO.Path]::GetFullPath($destination)
                    }
                )
            }
            continue
        }
        if ([System.IO.Path]::GetExtension($item.Name) -ine '.docx') {
            throw "Il file non e un DOCX: $inputPath"
        }

        if ($DestinationDirectory) {
            $destination = Join-Path $DestinationDirectory ([System.IO.Path]::ChangeExtension($item.Name, '.md'))
        }
        else {
            $destination = [System.IO.Path]::ChangeExtension($item.FullName, '.md')
        }
        $jobs.Add(
            [PSCustomObject]@{
                Source = $item.FullName
                Destination = [System.IO.Path]::GetFullPath($destination)
            }
        )
    }

    if ($jobs.Count -eq 0) {
        throw 'Non e stato trovato nessun file DOCX.'
    }
    $destinations = @{}
    foreach ($job in $jobs) {
        $key = $job.Destination.ToLowerInvariant()
        if ($destinations.ContainsKey($key)) {
            throw "Piu file DOCX genererebbero lo stesso file di output: $($job.Destination)"
        }
        $destinations[$key] = $true
    }
    return ,$jobs
}

function Invoke-Conversion {
    param(
        [PSCustomObject]$Job,
        [string]$PandocExecutable,
        [System.Collections.Generic.List[object]]$CorrectionsList,
        [bool]$Overwrite,
        [bool]$ExtractMedia
    )

    $mediaDirectory = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($Job.Destination),
        [System.IO.Path]::GetFileNameWithoutExtension($Job.Destination) + '_media'
    )
    if (-not $Overwrite -and (Test-Path -LiteralPath $Job.Destination)) {
        throw "Il file esiste gia: $($Job.Destination) (usa -Force per sovrascriverlo)"
    }
    if ($ExtractMedia -and -not $Overwrite -and (Test-Path -LiteralPath $mediaDirectory)) {
        throw "La cartella media esiste gia: $mediaDirectory (usa -Force per sovrascriverla)"
    }

    $destinationParent = [System.IO.Path]::GetDirectoryName($Job.Destination)
    [void][System.IO.Directory]::CreateDirectory($destinationParent)
    $pandocArguments = @('--from=docx', '--to=gfm', '--wrap=none')
    if ($ExtractMedia) {
        $pandocArguments += "--extract-media=$([System.IO.Path]::GetFileName($mediaDirectory))"
    }
    $pandocArguments += $Job.Source
    $result = Invoke-Pandoc $PandocExecutable $pandocArguments $destinationParent
    if ($result.ExitCode -ne 0) {
        $details = $result.StdErr.Trim()
        if (-not $details) {
            $details = "codice di uscita $($result.ExitCode)"
        }
        throw "Pandoc non ha convertito '$($Job.Source)': $details"
    }

    $cleaned = Clean-Markdown $result.StdOut $CorrectionsList
    $temporaryPath = Join-Path $destinationParent ([guid]::NewGuid().ToString('N') + '.md.tmp')
    try {
        [System.IO.File]::WriteAllText(
            $temporaryPath,
            $cleaned,
            [System.Text.UTF8Encoding]::new($false)
        )
        if (Test-Path -LiteralPath $Job.Destination) {
            Move-Item -LiteralPath $temporaryPath -Destination $Job.Destination -Force
        }
        else {
            Move-Item -LiteralPath $temporaryPath -Destination $Job.Destination
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

$pandoc = Resolve-Pandoc $PandocPath
$correctionsList = Read-Corrections $Corrections
$resolvedOutputDir = $null
if ($OutputDir) {
    $resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
}
$jobs = Get-ConversionJobs $Inputs $resolvedOutputDir ([bool]$Recursive)

$failed = $false
foreach ($job in $jobs) {
    try {
        Invoke-Conversion $job $pandoc $correctionsList ([bool]$Force) (-not [bool]$NoExtractMedia)
        Write-Output "Convertito: $($job.Source) -> $($job.Destination)"
    }
    catch {
        [Console]::Error.WriteLine("Errore: $($_.Exception.Message)")
        $failed = $true
    }
}

if ($failed) {
    exit 1
}
exit 0
