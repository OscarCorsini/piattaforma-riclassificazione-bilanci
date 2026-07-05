# Piattaforma di Riclassificazione Finanziaria

Applicazione Streamlit locale che estrae dati da situazioni contabili in PDF,
li riclassifica tramite Google Gemini e popola un template Excel di
stima del margine.

## File del progetto

- `app.py` — interfaccia Streamlit.
- `config.py` — categorie di riclassificazione, mapping celle Excel per anno, configurazione Gemini.
- `pdf_extractor.py` — estrazione testo dai PDF (pdfplumber + fallback PyPDF2).
- `gemini_client.py` — prompt e chiamata a Google Gemini, parsing risposta.
- `excel_writer.py` — scrittura dei valori nel template Excel con openpyxl (preserva formule e formattazioni).
- `requirements.txt` — dipendenze Python.

## Installazione

```bash
pip install -r requirements.txt
```

## Configurazione della chiave Gemini

Genera una chiave API gratuita su https://aistudio.google.com/apikey con il
tuo account Google (nessuna carta di credito richiesta per il piano
gratuito), poi impostala come variabile d'ambiente prima di avviare l'app:

```powershell
$env:GEMINI_API_KEY = "<la-tua-chiave>"
$env:GEMINI_MODEL   = "gemini-2.0-flash"   # opzionale, ha un default
```

Su macOS/Linux usa `export VAR=valore` al posto di `$env:VAR = "valore"`.

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

Nessuna carta di credito richiesta. Repository:
https://github.com/OscarCorsini/piattaforma-riclassificazione-bilanci
(già creato e pubblicato).

1. Vai su share.streamlit.io, accedi con l'account GitHub, clicca
   "New app" e seleziona il repository, branch `main`, file principale
   `app.py`.
2. Prima di avviare il deploy, apri "Advanced settings" -> "Secrets" e
   incolla il contenuto di `.streamlit/secrets.toml.example` con la tua
   chiave Gemini reale.
3. Clicca "Deploy". Dopo qualche minuto l'app è online con un link
   pubblico.
4. In "Settings" -> "Sharing" puoi restringere l'accesso solo a email
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
- Se Gemini restituisce una risposta non conforme al formato JSON
  atteso, l'app lo segnala senza scrivere dati incompleti nel file Excel.
