# Skill DOCX in Markdown per Windows

Questa skill converte documenti Word `.docx` in Markdown GitHub Flavored Markdown (GFM) con Pandoc. Dopo la conversione, il modello AI selezionato in Copilot Chat puo correggere refusi e residui di formattazione, mantenendo contenuto e significato originali.

## Requisiti

1. Windows 10 o 11.
2. Visual Studio Code.
3. GitHub Copilot disponibile in VS Code e accesso completato al proprio account.
4. Pandoc. L'installer puo installarlo tramite winget.

Python non e necessario.

## Installazione rapida

1. Apri la [repository della skill](https://github.com/AndreaSalmaso/docx-to-markdown).
2. Fai clic su **Code**, poi **Download ZIP**. Estrai lo ZIP in una cartella.
3. Apri la cartella estratta in Esplora file. Fai clic sulla barra del percorso (oppure premi `Ctrl+L`), scrivi `powershell` e premi **Invio**.
4. Nella finestra PowerShell esegui:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
   ```

5. Conferma le operazioni quando richiesto. Se Pandoc non e installato, lo script propone di installarlo con winget.
6. Se VS Code era gia aperto durante l'installazione, chiudilo e riaprilo affinche rilevi la skill e Pandoc.

L'installer copia la skill in `%USERPROFILE%\.copilot\skills\docx-to-markdown` e crea le cartelle necessarie. Se la skill esiste gia, chiede conferma prima di aggiornarla. Se winget non e disponibile, segui la procedura manuale qui sotto.

## Installazione manuale

### 1. Installa Pandoc

Apri PowerShell e incolla:

```powershell
winget install --id JohnMacFarlane.Pandoc -e
```

Quando l'installazione termina, chiudi e riapri PowerShell e VS Code: le finestre gia aperte potrebbero non avere ancora aggiornato il `PATH`.

### 2. Verifica Pandoc

In una nuova finestra PowerShell esegui:

```powershell
pandoc --version
```

Se Windows non trova `pandoc`, controlla che la cartella contenente `pandoc.exe` (per esempio `C:\Program Files\Pandoc` o `%LOCALAPPDATA%\Pandoc`) sia nel `PATH` utente:

1. Dal menu **Start**, cerca **Modifica le variabili di ambiente per l'account**.
2. Seleziona **Path** nelle variabili utente e scegli **Modifica**.
3. Scegli **Nuovo** e inserisci la cartella che contiene `pandoc.exe`, non il percorso del file stesso.
4. Conferma le finestre, poi riapri PowerShell e VS Code.
5. Riprova `pandoc --version`.

### 3. Installa la skill a mano

Copia la cartella `docx-to-markdown` estratta in:

```text
%USERPROFILE%\.copilot\skills
```

La struttura risultante deve essere `%USERPROFILE%\.copilot\skills\docx-to-markdown\SKILL.md`. Se le cartelle non esistono, creale. Per installarla solo in un progetto, copia la skill in `<cartella-del-progetto>\.github\skills\docx-to-markdown`.

Chiudi e riapri VS Code e controlla che `/docx-to-markdown` compaia in Copilot Chat.

## Uso

In Copilot Chat, seleziona `/docx-to-markdown` e indica un file DOCX o una cartella. Per esempio:

```text
/docx-to-markdown "C:\Documenti\manuale.docx"
```

I file Markdown vengono salvati accanto ai documenti. Le immagini e gli altri media estratti restano nella cartella `<nome-documento>_media`. I formati non visualizzati dal browser, come EMF, WMF e TIFF, vengono convertiti in PNG mantenendo il nome e il numero originali (per esempio `image2.emf` diventa `image2.png`).

Per impostazione predefinita, lo script non sovrascrive un Markdown o una cartella media gia presenti. Usa `-Force` solo se vuoi sostituire i risultati esistenti.

## Test

Con Pandoc installato, apri PowerShell nella cartella della skill ed esegui:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\test_docx_to_markdown.ps1"
```

I test creano documenti temporanei e li eliminano alla fine.

## Licenza

Questa repository non include una licenza. Il codice e visibile pubblicamente, ma la pubblicazione non concede automaticamente il permesso di copiarlo, modificarlo o ridistribuirlo.
