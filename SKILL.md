---
name: docx-to-markdown
description: 'Convert DOCX files to GitHub Flavored Markdown with Pandoc and Windows PowerShell, rasterize unsupported EMF/WMF/TIFF images to PNG, then review and clean up the Markdown while preserving meaning. Use for Word-to-Markdown conversion and DOCX proofreading.'
argument-hint: '[file.docx o cartella] [altri file o cartelle]'
user-invocable: true
disable-model-invocation: true
---

# Convertire DOCX in Markdown

Usa questa skill quando l’utente chiede di convertire documenti Word in Markdown o di revisionare il testo dopo la conversione.

## Procedura

1. Interpreta i percorsi o i file DOCX indicati dopo il comando. Se non è indicato alcun percorso, cerca i file `.docx` nella cartella di lavoro attiva. Includi le sottocartelle solo se richiesto.
2. Controlla che i file esistano e che Windows PowerShell sia disponibile. Risolvi Pandoc prima cercando `pandoc` nel `PATH`; se non è disponibile, controlla `%LOCALAPPDATA%\Pandoc\pandoc.exe`, `C:\Program Files\Pandoc\pandoc.exe` e `C:\Program Files (x86)\Pandoc\pandoc.exe`. Se nessuna posizione contiene Pandoc, spiega che è necessario e fermati senza simulare la conversione. Se estrai media, verifica anche `System.Drawing`. Non installare software senza consenso esplicito, non inviare documenti a servizi esterni.
3. Esegui lo script `scripts/docx_to_markdown.ps1` incluso in questa skill, individuando la cartella della skill a partire da questo `SKILL.md`. Non usare uno script omonimo trovato nel workspace. Passa sempre a `-PandocPath` il percorso risolto:

   ```powershell
   $pandocPath = $null
   $pandocCommand = Get-Command pandoc -ErrorAction SilentlyContinue
   if ($pandocCommand) { $pandocPath = $pandocCommand.Source }
   if (-not $pandocPath -and $env:LOCALAPPDATA) {
       $candidate = Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
       if (Test-Path -LiteralPath $candidate) { $pandocPath = $candidate }
   }
   if (-not $pandocPath -and $env:ProgramFiles) {
       $candidate = Join-Path $env:ProgramFiles 'Pandoc\pandoc.exe'
       if (Test-Path -LiteralPath $candidate) { $pandocPath = $candidate }
   }
   if (-not $pandocPath -and ${env:ProgramFiles(x86)}) {
       $candidate = Join-Path ${env:ProgramFiles(x86)} 'Pandoc\pandoc.exe'
       if (Test-Path -LiteralPath $candidate) { $pandocPath = $candidate }
   }
   if (-not $pandocPath) { throw 'Pandoc non trovato: installalo con winget.' }
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<cartella-della-skill>\scripts\docx_to_markdown.ps1" "PERCORSO\documento.docx" -PandocPath $pandocPath
   ```

   Azzera `$pandocPath` prima di risolverlo se riusi la sessione. Passa tutti i percorsi DOCX come argomenti distinti. Per una cartella, lo script converte i DOCX direttamente contenuti; aggiungi `-Recursive` solo se richiesto. Per scegliere la destinazione usa `-OutputDir "PERCORSO\destinazione"`. Usa `-NoExtractMedia` solo se l’utente chiede di non estrarre le immagini; usa `-Force` solo con autorizzazione esplicita a sovrascrivere.
4. Per ogni conversione con estrazione media attiva (impostazione predefinita), esegui `scripts/convert_unsupported_images.ps1` incluso in questa skill una sola volta per ciascun Markdown, passando:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<cartella-della-skill>\scripts\convert_unsupported_images.ps1" -MarkdownPath "PERCORSO\documento.md" -MediaDirectory "PERCORSO\documento_media\media"
   ```

   Se lo script principale è stato avviato con `-OutputDir`, la cartella media è `<OutputDir>\<nome-documento>_media\media`. Non eseguire questo passaggio se è stato usato `-NoExtractMedia`.
   
   Questo helper richiede Windows PowerShell e `System.Drawing`, già disponibile su Windows. Se la conversione o l’aggiornamento dei riferimenti non riesce, mostra l’errore e non dichiarare completato il documento.
5. Se estrai media, converti i formati non visualizzati dal browser come EMF/WMF/TIFF con l’helper indicato sopra; conserva JPEG, PNG, GIF, WebP, BMP e SVG. Mantieni il nome base e il numero originali (`image2.emf` → `image2.png`), riusa un PNG omonimo valido e rimuovi la sorgente non supportata solo dopo il salvataggio dei link aggiornati.
6. Dopo la conversione meccanica, revisiona il Markdown con il modello LLM attivo in Copilot Chat. Correggi refusi e residui di conversione, verifica titoli, tabelle ed elenchi GFM e preserva significato, contenuti, link e riferimenti ai media. Non inventare contenuti. Lo script applica solo pulizie deterministiche semplici e non sostituisce questa revisione.
7. Se conversione, rasterizzazione o aggiornamento dei link falliscono, mostra l’errore e non dichiarare completato il documento. Verifica che il Markdown resti GFM valido e che le immagini collegate esistano.
