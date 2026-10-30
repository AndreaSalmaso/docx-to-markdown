param(
    [Parameter(Mandatory = $true)]
    [string]$MarkdownPath,

    [Parameter(Mandatory = $true)]
    [string]$MediaDirectory
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$markdownFullPath = [System.IO.Path]::GetFullPath($MarkdownPath)
$mediaFullPath = [System.IO.Path]::GetFullPath($MediaDirectory)
if (-not (Test-Path -LiteralPath $markdownFullPath -PathType Leaf)) {
    throw "File Markdown non trovato: $markdownFullPath"
}
if (-not (Test-Path -LiteralPath $mediaFullPath -PathType Container)) {
    throw "Cartella media non trovata: $mediaFullPath"
}

$markdownRoot = [System.IO.Path]::GetDirectoryName($markdownFullPath)
$unsupportedExtensions = @('.emf', '.tif', '.tiff', '.wmf')
$imagePattern = '!\[(?<alt>[^\]]*)\]\((?<target><[^>]+>|[^)\s]+)(?<suffix>\s+(?:"[^"]*"|''[^'']*''|\([^)]*\)))?\)'
$markdown = [System.IO.File]::ReadAllText($markdownFullPath, [System.Text.Encoding]::UTF8)
$searchCharacters = $markdown.ToCharArray()
$insideFence = $false
$fenceCharacter = $null
$fenceLength = 0
$lineStart = 0
while ($lineStart -lt $markdown.Length) {
    $lineEnd = $markdown.IndexOf("`n", $lineStart)
    if ($lineEnd -lt 0) {
        $lineEnd = $markdown.Length
    }
    $line = $markdown.Substring($lineStart, $lineEnd - $lineStart).TrimEnd("`r")
    $fence = [regex]::Match($line, '^ {0,3}(`{3,}|~{3,})')

    if ($insideFence -or $fence.Success) {
        for ($index = $lineStart; $index -lt $lineEnd; $index++) {
            if ($searchCharacters[$index] -ne "`r") {
                $searchCharacters[$index] = ' '
            }
        }
    }
    if ($insideFence) {
        $closingFence = [regex]::Match($line, '^ {0,3}(`+|~+)[ \t]*$')
        if (
            $closingFence.Success -and
            $closingFence.Groups[1].Value[0] -eq $fenceCharacter -and
            $closingFence.Groups[1].Length -ge $fenceLength
        ) {
            $insideFence = $false
            $fenceCharacter = $null
            $fenceLength = 0
        }
    }
    elseif ($fence.Success) {
        $insideFence = $true
        $fenceCharacter = $fence.Groups[1].Value[0]
        $fenceLength = $fence.Groups[1].Length
    }

    $lineStart = $lineEnd + 1
}
$searchMarkdown = -join $searchCharacters
$matches = [regex]::Matches($searchMarkdown, $imagePattern)
$references = @()
$mediaPrefix = $mediaFullPath.TrimEnd('\') + '\'
$markdownRootUri = New-Object System.Uri($markdownRoot.TrimEnd('\') + '\')

foreach ($match in $matches) {
    $target = $match.Groups['target'].Value.Trim('<', '>')
    if ($target -match '^[a-zA-Z][a-zA-Z0-9+.-]*:' -and $target -notmatch '^[a-zA-Z]:[\\/]') {
        continue
    }

    $targetPath = ($target -split '[?#]', 2)[0]
    $unescapedTarget = [System.Uri]::UnescapeDataString($targetPath)
    if ([string]::IsNullOrWhiteSpace($unescapedTarget)) {
        continue
    }
    if (-not [System.IO.Path]::IsPathRooted($unescapedTarget)) {
        $referencePath = [System.IO.Path]::GetFullPath(
            [System.IO.Path]::Combine($markdownRoot, $unescapedTarget.Replace('/', '\'))
        )
    }
    else {
        $referencePath = [System.IO.Path]::GetFullPath($unescapedTarget)
    }

    if (-not $referencePath.StartsWith($mediaPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        continue
    }
    if (-not (Test-Path -LiteralPath $referencePath -PathType Leaf)) {
        continue
    }

    $extension = [System.IO.Path]::GetExtension($referencePath).ToLowerInvariant()
    if ($extension -notin $unsupportedExtensions) {
        continue
    }

    $references += [PSCustomObject]@{
        Match = $match
        SourcePath = $referencePath
    }
}

$converted = @{}
$createdFiles = @()
$sourcesToDelete = @{}
$rewrites = @()
$destinationSources = @{}
foreach ($reference in $references) {
    $sourcePath = $reference.SourcePath
    if (-not $converted.ContainsKey($sourcePath)) {
        $destinationName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath) + '.png'
        $destinationPath = Join-Path $mediaFullPath $destinationName
        $destinationKey = $destinationPath.ToLowerInvariant()
        if ($destinationSources.ContainsKey($destinationKey)) {
            foreach ($createdPath in $createdFiles) {
                if (Test-Path -LiteralPath $createdPath) {
                    Remove-Item -LiteralPath $createdPath -Force
                }
            }
            throw "Piu immagini non supportate avrebbero lo stesso PNG di destinazione: '$destinationPath'."
        }
        $destinationSources[$destinationKey] = $sourcePath

        if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
            $existingImage = $null
            try {
                $existingImage = [System.Drawing.Image]::FromFile($destinationPath)
            }
            catch {
                throw "Il PNG di destinazione esiste ma non e leggibile; non lo sovrascrivo: '$destinationPath'."
            }
            finally {
                if ($existingImage) { $existingImage.Dispose() }
            }
        }
        else {
            $sourceImage = $null
            $bitmap = $null
            $graphics = $null
            try {
                $sourceImage = [System.Drawing.Image]::FromFile($sourcePath)
                $bitmap = [System.Drawing.Bitmap]::new(
                    $sourceImage.Width,
                    $sourceImage.Height,
                    [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
                )
                $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                $graphics.Clear([System.Drawing.Color]::Transparent)
                $graphics.DrawImage($sourceImage, 0, 0, $sourceImage.Width, $sourceImage.Height)
                $bitmap.Save($destinationPath, [System.Drawing.Imaging.ImageFormat]::Png)
                $createdFiles += $destinationPath
            }
            catch {
                if (Test-Path -LiteralPath $destinationPath) {
                    Remove-Item -LiteralPath $destinationPath -Force
                }
                foreach ($createdPath in $createdFiles) {
                    if (Test-Path -LiteralPath $createdPath) {
                        Remove-Item -LiteralPath $createdPath -Force
                    }
                }
                throw "Impossibile convertire '$sourcePath' in PNG: $($_.Exception.Message)"
            }
            finally {
                if ($graphics) { $graphics.Dispose() }
                if ($bitmap) { $bitmap.Dispose() }
                if ($sourceImage) { $sourceImage.Dispose() }
            }
        }
        $converted[$sourcePath] = $destinationPath
        $sourcesToDelete[$sourcePath] = $true
    }

    $markdownTarget = $markdownRootUri.MakeRelativeUri(
        (New-Object System.Uri($converted[$sourcePath]))
    ).ToString()
    $markdownTarget = $markdownTarget.Replace('\', '/')
    $replacement = '![' + $reference.Match.Groups['alt'].Value + '](' +
        $markdownTarget + $reference.Match.Groups['suffix'].Value + ')'
    $rewrites += [PSCustomObject]@{
        Index = $reference.Match.Index
        Length = $reference.Match.Length
        Replacement = $replacement
    }
}

foreach ($rewrite in ($rewrites | Sort-Object Index -Descending)) {
    $markdown = $markdown.Remove($rewrite.Index, $rewrite.Length).Insert(
        $rewrite.Index,
        $rewrite.Replacement
    )
}

if ($references.Count -gt 0) {
    try {
        [System.IO.File]::WriteAllText(
            $markdownFullPath,
            $markdown,
            [System.Text.UTF8Encoding]::new($false)
        )
    }
    catch {
        foreach ($createdPath in $createdFiles) {
            if (Test-Path -LiteralPath $createdPath) {
                Remove-Item -LiteralPath $createdPath -Force
            }
        }
        throw "Impossibile aggiornare il Markdown '$markdownFullPath': $($_.Exception.Message)"
    }
}

foreach ($sourcePath in $sourcesToDelete.Keys) {
    try {
        Remove-Item -LiteralPath $sourcePath -Force
    }
    catch {
        throw "Markdown aggiornato, ma non e stato possibile rimuovere l'immagine originale '$sourcePath': $($_.Exception.Message)"
    }
}

foreach ($destinationPath in $converted.Values) {
    Write-Output "Immagine disponibile in PNG: $destinationPath"
}
