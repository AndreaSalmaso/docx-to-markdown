[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$Destination = (Join-Path $HOME '.copilot\skills\docx-to-markdown')
)

$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
    throw 'Questo installer supporta solo Windows.'
}

$requiredFiles = @(
    (Join-Path $PSScriptRoot 'SKILL.md'),
    (Join-Path $PSScriptRoot 'README.md'),
    (Join-Path $PSScriptRoot 'scripts\docx_to_markdown.ps1'),
    (Join-Path $PSScriptRoot 'scripts\convert_unsupported_images.ps1')
)
foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Il pacchetto della skill e incompleto; file mancante: $requiredFile"
    }
}

function Find-Pandoc {
    $command = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }
    foreach ($candidate in @(
        $(if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe' }),
        $(if ($env:ProgramFiles) { Join-Path $env:ProgramFiles 'Pandoc\pandoc.exe' }),
        $(if (${env:ProgramFiles(x86)}) { Join-Path ${env:ProgramFiles(x86)} 'Pandoc\pandoc.exe' })
    )) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    return $null
}

$pandocPath = Find-Pandoc
if (-not $pandocPath) {
    $wingetCommand = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $wingetCommand) {
        throw 'Pandoc non e installato e winget non e disponibile. Installa prima App Installer dal Microsoft Store oppure segui la guida manuale nel README.md.'
    }
    if ($PSCmdlet.ShouldProcess('Pandoc', 'Install via winget')) {
        & $wingetCommand.Source install --id JohnMacFarlane.Pandoc -e `
            --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) {
            throw "L'installazione di Pandoc con winget e fallita (codice $LASTEXITCODE)."
        }
    }
    else {
        Write-Output 'Installazione di Pandoc non eseguita.'
        return
    }

    $pandocPath = Find-Pandoc
    if (-not $pandocPath) {
        throw 'Pandoc risulta installato, ma non e ancora disponibile. Chiudi e riapri PowerShell e VS Code, poi verifica pandoc --version.'
    }
}

$pandocVersion = & $pandocPath --version 2>&1
if ($LASTEXITCODE -ne 0 -or -not $pandocVersion) {
    throw "Pandoc non si avvia correttamente dal percorso: $pandocPath"
}

$pandocDirectory = Split-Path -Parent $pandocPath
$processPathContainsPandoc = $false
foreach ($pathEntry in @($env:PATH -split ';' | Where-Object { $_ })) {
    try {
        if ([System.IO.Path]::GetFullPath($pathEntry).TrimEnd('\') -ieq $pandocDirectory.TrimEnd('\')) {
            $processPathContainsPandoc = $true
            break
        }
    }
    catch {
        continue
    }
}
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$userPathEntries = @($userPath -split ';' | Where-Object { $_ })
$userPathContainsPandoc = $false
foreach ($pathEntry in $userPathEntries) {
    try {
        if ([System.IO.Path]::GetFullPath($pathEntry).TrimEnd('\') -ieq $pandocDirectory.TrimEnd('\')) {
            $userPathContainsPandoc = $true
            break
        }
    }
    catch {
        continue
    }
}
$pathConfigured = $processPathContainsPandoc -or $userPathContainsPandoc
if (
    -not $processPathContainsPandoc -and
    -not $userPathContainsPandoc -and
    $PSCmdlet.ShouldProcess('User PATH', "Add $pandocDirectory")
) {
    [Environment]::SetEnvironmentVariable(
        'Path',
        (@($userPathEntries + $pandocDirectory) -join ';'),
        'User'
    )
    $env:PATH = $env:PATH.TrimEnd(';') + ';' + $pandocDirectory
    $pathConfigured = $true
}

$sourcePath = [System.IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$destinationPath = [System.IO.Path]::GetFullPath($Destination).TrimEnd('\')
if ($sourcePath -ieq $destinationPath) {
    Write-Output "La skill e gia nella cartella di installazione: $destinationPath"
    Write-Output "Pandoc disponibile: $pandocPath"
    return
}

$skillInstalled = $false
if ($PSCmdlet.ShouldProcess($destinationPath, 'Install or update the personal skill')) {
    [void][System.IO.Directory]::CreateDirectory($destinationPath)
    Copy-Item -Path (Join-Path $PSScriptRoot '*') -Destination $destinationPath -Recurse -Force

    foreach ($obsoleteScript in @(
        (Join-Path $destinationPath 'scripts\docx_to_markdown.py'),
        (Join-Path $destinationPath 'scripts\test_docx_to_markdown.py')
    )) {
        if (Test-Path -LiteralPath $obsoleteScript -PathType Leaf) {
            Remove-Item -LiteralPath $obsoleteScript -Force
        }
    }
    $skillInstalled = $true
}

if ($skillInstalled) {
    Write-Output "Skill installata in: $destinationPath"
}
elseif ($WhatIfPreference) {
    Write-Output "WhatIf: la skill non e stata copiata in $destinationPath."
}
else {
    Write-Output "Installazione della skill annullata: $destinationPath"
}
Write-Output "Pandoc disponibile: $pandocPath"
if (-not $pathConfigured) {
    Write-Output "Pandoc non e stato aggiunto al PATH. Puoi usarlo dal percorso: $pandocPath"
}
Write-Output 'Per usare il comando slash, VS Code deve avere GitHub Copilot attivo e un account con accesso completato.'
Write-Output 'Riapri PowerShell e VS Code per aggiornare il PATH e caricare la skill.'
