# Piattaforma di Riclassificazione Finanziaria

Applicazione Streamlit locale che estrae dati da situazioni contabili in PDF,
li riclassifica tramite Microsoft Copilot e popola un template Excel di
stima del margine.

## File del progetto

- `app.py` — interfaccia Streamlit.
- `config.py` — categorie di riclassificazione, mapping celle Excel per anno, configurazione Copilot.
- `pdf_extractor.py` — estrazione testo dai PDF (pdfplumber + fallback PyPDF2).
- `copilot_client.py` — prompt e chiamata a Microsoft Copilot / Azure OpenAI, parsing risposta.
- `excel_writer.py` — scrittura dei valori nel template Excel con openpyxl (preserva formule e formattazioni).
- `requirements.txt` — dipendenze Python.

## Installazione

```bash
pip install -r requirements.txt
```

## Configurazione credenziali Copilot

Prima di avviare l'app, imposta le variabili d'ambiente con i dati del tuo
endpoint aziendale (Azure OpenAI / Copilot Studio dietro Azure AI Foundry):

```powershell
$env:COPILOT_ENDPOINT   = "https://<tuo-endpoint>.openai.azure.com/"
$env:COPILOT_API_KEY    = "<la-tua-chiave>"
$env:COPILOT_DEPLOYMENT = "<nome-deployment>"     # es. gpt-4o
$env:COPILOT_API_VERSION = "2024-06-01"           # opzionale, ha un default
```

Su macOS/Linux usa `export VAR=valore` al posto di `$env:VAR = "valore"`.

Se la tua organizzazione espone Copilot tramite un connettore diverso da
Azure OpenAI (es. Microsoft Graph Copilot API o un gateway REST interno),
apri `copilot_client.py` e sostituisci il corpo della funzione `_call_model()`
con la chiamata equivalente, mantenendo il resto invariato.

## Adattare il mapping delle celle Excel

Il file `config.py` contiene `CELL_MAPPING`: un dizionario di esempio che
associa ogni categoria (es. "Corrispettivi normali") a una cella del
template ("C6", "D6", ...) per ciascun anno. Apri il tuo `Template.xlsx`
reale, individua le celle corrette e aggiorna questo dizionario di
conseguenza. Per aggiungere un nuovo anno, duplica un blocco esistente e
cambia le colonne.

## Avvio locale dell'applicazione

```bash
streamlit run app.py
```

L'app si apre nel browser all'indirizzo indicato in console (di norma
http://localhost:8501).

## Pubblicazione su web (Streamlit Community Cloud, gratuito)

Nessuna carta di credito richiesta. Serve solo l'account GitHub
**OscarCorsini** già esistente.

1. Su github.com, crea un nuovo repository (es. "riclassificazione-bilanci"),
   privato o pubblico a scelta.
2. Dal tuo computer, nella cartella del progetto:
   ```bash
   git init
   git add .
   git commit -m "Prima versione piattaforma riclassificazione"
   git branch -M main
   git remote add origin https://github.com/OscarCorsini/riclassificazione-bilanci.git
   git push -u origin main
   ```
   (il file `.gitignore` già presente esclude automaticamente segreti e
   file temporanei dal caricamento)
3. Vai su share.streamlit.io, accedi con l'account GitHub, clicca
   "New app" e seleziona il repository appena creato, branch `main`,
   file principale `app.py`.
4. Prima di avviare il deploy, apri "Advanced settings" -> "Secrets" e
   incolla il contenuto di `.streamlit/secrets.toml.example` con i tuoi
   valori reali (endpoint e chiave Copilot).
5. Clicca "Deploy". Dopo qualche minuto l'app è online con un link
   pubblico (es. `https://riclassificazione-bilanci.streamlit.app`).
6. In "Settings" -> "Sharing" puoi restringere l'accesso solo a email
   specifiche, se l'app non deve essere pubblica.

## Flusso d'uso

1. Carica il template Excel.
2. Seleziona l'anno di riferimento.
3. Carica uno o più PDF della situazione contabile.
4. Premi "Avvia elaborazione".
5. Scarica il file Excel completato con i dati riclassificati.

## Note

- Il file template caricato non viene mai modificato sul disco: l'app
  lavora su una copia in memoria e restituisce il risultato in download.
- Se un PDF è una scansione senza testo selezionabile (nessun OCR), l'app
  mostra un avviso chiaro invece di generare un output vuoto.
- Se Copilot restituisce una risposta non conforme al formato JSON
  atteso, l'app lo segnala senza scrivere dati incompleti nel file Excel.
