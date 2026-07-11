# -*- coding: utf-8 -*-
"""
groq_client.py
==============
Modulo di integrazione con Groq (motore AI usato per la riclassificazione
contabile). Groq offre un'API compatibile con lo standard OpenAI, gratuita
e senza richiesta di carta di credito.

APPROCCIO (importante):
------------------------
In una prima versione, si chiedeva al modello di leggere l'intero bilancio
e contemporaneamente sommare/classificare le voci in 20+ macro-categorie in
un solo colpo. Si e' rivelato inaffidabile: il modello tendeva ad accorpare
le voci sotto poche categorie generiche, anche quando il documento le
elencava chiaramente in modo distinto (es. "ENERGIA ELETTRICA",
"MERCI C/ACQUISTI - CARBURANTI" come righe separate con proprio codice
conto e proprio importo).

In una seconda versione, si chiedeva all'AI di estrarre solo le voci e di
giudicare DA SOLA quali fossero subtotali gerarchici da scartare (es. "3.66"
e' il subtotale di "3.66.05.001" + "3.66.05.002"). Anche questo si e'
rivelato poco affidabile: il modello a volte estraeva ancora la riga
aggregata (es. "MERCI C/ACQUISTI" o "COSTI PER SERVIZI") invece delle righe
di dettaglio sottostanti, con il risultato che gli importi finivano di
nuovo nelle categorie generiche invece che in quelle specifiche.

L'approccio attuale toglie ANCHE questo giudizio all'AI: il modello si
limita a estrarre OGNI riga del documento (comprese le righe aggregate),
insieme al proprio codice conto quando presente - un compito puramente
meccanico. La deduplicazione gerarchica (capire quali codici sono
subtotali di altri codici piu' specifici e scartarli) viene fatta DOPO, in
modo matematico/deterministico, confrontando i codici conto estratti
(classificatore.py, funzione filtra_subtotali_gerarchia). La classificazione
nella categoria corretta viene infine svolta anch'essa in modo
deterministico da classificatore.py (parole chiave, puro codice Python).
In questo modo l'AI non deve piu' prendere alcuna decisione delicata: deve
solo leggere e trascrivere fedelmente, il che e' un compito molto piu'
alla sua portata.

STRUTTURA DEL RISULTATO:
-------------------------
RiclassificazioneResult tiene, per ciascuna categoria (entrate/uscite/
gestione fiscale), la LISTA delle singole voci originali che vi sono state
assegnate (oggetti Voce: descrizione + importo), non solo il totale. Questo
permette sia di calcolare i totali per categoria (metodi totale_entrate(),
totale_uscite(), totale_gestione_fiscale(), usati per scrivere le celle
riassuntive nel template Excel) sia di generare un foglio di dettaglio con
ogni singola voce originale (utile per verificare rapidamente che la
classificazione sia corretta su un bilancio reale).

Questo modulo quindi:
  1. Costruisce un prompt che chiede l'estrazione strutturata di ogni voce
     (inclusi i codici conto, quando presenti).
  2. Invia il testo estratto dai PDF a Groq chiedendo una risposta in
     formato JSON rigido.
  3. Valida e normalizza la risposta prima di restituirla al chiamante.

>>> DOVE INSERIRE LA CHIAVE API <<<
La chiave (GROQ_API_KEY) e' letta da config.py (classe GroqConfig), che a
sua volta la legge da st.secrets (Streamlit Community Cloud) o da
variabile d'ambiente (esecuzione locale).

Per generarla: vai su https://console.groq.com/keys con un account email
o Google, clicca "Create API Key" e copiala. Nessuna carta di credito
richiesta, piano gratuito permanente.
"""

import json
import re
from dataclasses import dataclass, field
from typing import Dict, List

from config import GROQ_CONFIG, GESTIONE_FISCALE_CATEGORIE


class GroqConfigError(Exception):
    """Chiave API mancante o non valida."""


class GroqResponseError(Exception):
    """La risposta di Groq non e' nel formato JSON atteso."""


@dataclass
class Voce:
    """Una singola voce originale del bilancio (descrizione + importo),
    gia' assegnata a una categoria. Serve per mantenere il dettaglio
    verificabile oltre al semplice totale per categoria."""
    descrizione: str
    importo: float


@dataclass
class RiclassificazioneResult:
    # Categoria -> lista delle voci originali assegnate a quella categoria.
    entrate: Dict[str, List[Voce]] = field(default_factory=dict)
    uscite: Dict[str, List[Voce]] = field(default_factory=dict)
    gestione_fiscale: Dict[str, List[Voce]] = field(default_factory=dict)
    note: List[str] = field(default_factory=list)

    def totale_entrate(self) -> Dict[str, float]:
        return {cat: sum(v.importo for v in voci) for cat, voci in self.entrate.items()}

    def totale_uscite(self) -> Dict[str, float]:
        return {cat: sum(v.importo for v in voci) for cat, voci in self.uscite.items()}

    def totale_gestione_fiscale(self) -> Dict[str, float]:
        return {cat: sum(v.importo for v in voci) for cat, voci in self.gestione_fiscale.items()}


# ---------------------------------------------------------------------------
# COSTRUZIONE DEL PROMPT (SOLO ESTRAZIONE, NESSUNA CLASSIFICAZIONE)
# ---------------------------------------------------------------------------
def _build_system_prompt() -> str:
    fiscale = "\n".join(f"  - {c}" for c in GESTIONE_FISCALE_CATEGORIE)

    return f"""Sei un assistente esperto di contabilita' agraria e bilanci d'esercizio.
Il tuo compito e' leggere una situazione contabile (bilancio di verifica,
mastrini, estratto conto economico) fornita come testo estratto da PDF ed
ESTRARRE OGNI SINGOLA VOCE elencata, senza classificarla ne' raggrupparla.

Non devi decidere in quale macro-categoria contabile mettere le voci, e NON
devi decidere quali righe sono "subtotali da scartare": queste due
decisioni le prende un altro sistema automatico, in modo matematico, dopo
la tua estrazione. Il tuo UNICO compito e' trascrivere fedelmente e
COMPLETAMENTE ogni riga del documento che ha un codice conto e un importo,
cosi' come appare, senza fare alcun giudizio su cosa tenere o scartare.

REGOLE DI ESTRAZIONE:
1. Analizza OGNI riga del bilancio di verifica/mastrino, sia nella sezione
   RICAVI/ENTRATE sia nella sezione COSTI/USCITE. I documenti hanno spesso
   un codice conto (es. "66/25/010" oppure "3.66.05.001") seguito da una
   descrizione abbreviata in maiuscolo e uno o piu' importi.
2. ESTRAI ASSOLUTAMENTE TUTTE le righe con un codice conto e un importo,
   COMPRESE le righe che ti sembrano "totali di un sottogruppo" (es.
   "3.66" COSTI P/MAT. PRI. 342.546,13, oppure "MERCI C/ACQUISTI"
   892.818,54): NON provare a indovinare se una riga e' un subtotale di
   altre righe piu' sotto - estraile TUTTE, sia quella piu' generale sia
   quelle piu' specifiche sotto di essa. Riporta sempre il codice conto
   esatto nel campo "codice_conto" (stringa, es. "3.66.05.001" o
   "66/25/010"): e' quello che permette al sistema automatico di capire
   quali righe sono duplicate/subtotali e scartarle correttamente al posto
   tuo. Se una riga non ha alcun codice conto visibile, lascia
   "codice_conto" vuoto ("").
3. Le UNICHE righe da NON estrarre sono quelle testuali di puro riepilogo
   senza un proprio codice conto specifico, tipicamente con etichette come
   "TOTALE", "PROGRESSIVO", "TOTALE RICAVI", "TOTALE COSTI", "UTILE
   D'ESERCIZIO", "PERDITA D'ESERCIZIO", "REDDITO IMPONIBILE", "A
   PAREGGIO": queste sono righe di chiusura del bilancio, non voci di
   costo/ricavo.
4. Non sommare mai insieme voci diverse: ogni riga distinta del documento
   e' una voce a se stante nel JSON, anche se il suo importo e' piccolo o
   sembra ripetersi.
5. Per ogni voce riporta il testo della descrizione ESATTAMENTE come
   appare nel documento (mantieni abbreviazioni, maiuscole), il codice
   conto nel campo "codice_conto", l'importo come numero (positivo, senza
   simboli di valuta ne' separatori delle migliaia) e "tipo": "entrata" se
   la voce e' un ricavo/provento, "uscita" se e' un costo/onere.
6. Estrai i dati di gestione fiscale (IVA vendite VE26, IVA acquisti VF27,
   imposta dovuta VL3, netto gestione) SOLO se esplicitamente presenti nel
   documento, nelle categorie:
{fiscale}
   Se non presenti, restituisci 0.0 per quella categoria.

FORMATO DI OUTPUT (OBBLIGATORIO):
Rispondi ESCLUSIVAMENTE con un oggetto JSON valido, senza testo aggiuntivo,
commenti, markdown o backtick, con questa struttura esatta:

{{
  "voci": [
    {{"descrizione": "<testo esatto della voce>", "codice_conto": "<codice o vuoto>", "importo": <numero>, "tipo": "entrata"}},
    {{"descrizione": "<testo esatto della voce>", "codice_conto": "<codice o vuoto>", "importo": <numero>, "tipo": "uscita"}}
  ],
  "gestione_fiscale": {{ "<categoria>": <valore_numerico>, ... }},
  "note": ["<eventuali osservazioni su righe ambigue, illeggibili o incerte>"]
}}
"""


def _build_user_prompt(anno: str, documenti_testo: List[str]) -> str:
    testo_concatenato = "\n\n----- NUOVO DOCUMENTO -----\n\n".join(documenti_testo)
    return f"""Anno di riferimento: {anno}

Testo estratto dai documenti contabili (uno o piu' PDF concatenati):

{testo_concatenato}

Estrai TUTTE le singole voci presenti nel testo sopra, secondo le
istruzioni di sistema. Non classificarle in categorie: limitati a
estrarre descrizione, importo e tipo (entrata/uscita) di ciascuna voce."""


# ---------------------------------------------------------------------------
# CHIAMATA AL MODELLO
# ---------------------------------------------------------------------------
def _call_model(system_prompt: str, user_prompt: str) -> str:
    """
    Effettua la chiamata all'API di Groq (compatibile OpenAI) e restituisce
    il contenuto testuale grezzo della risposta.
    """
    if not GROQ_CONFIG.api_key:
        raise GroqConfigError(
            "Chiave API di Groq non configurata. Genera una chiave gratuita su "
            "https://console.groq.com/keys e impostala come GROQ_API_KEY "
            "(vedi commenti in config.py / groq_client.py) prima di avviare "
            "l'elaborazione."
        )

    try:
        from openai import OpenAI, APIConnectionError
    except ImportError as exc:
        raise GroqConfigError(
            "Libreria 'openai' non installata. Esegui: pip install openai"
        ) from exc

    import httpx

    # Client HTTP dedicato:
    #   - trust_env=False: ignora eventuali variabili d'ambiente HTTP_PROXY/
    #     HTTPS_PROXY impostate dalla piattaforma di hosting (alcuni host
    #     cloud le settano per servizi interni e questo puo' interferire con
    #     la chiamata verso api.groq.com, causando un generico "Connection
    #     error" senza dettagli).
    #   - timeout di connessione separato e piu' breve del timeout totale,
    #     cosi' un problema di rete si manifesta rapidamente invece di
    #     attendere l'intero timeout_seconds prima di fallire.
    http_client = httpx.Client(
        trust_env=False,
        timeout=httpx.Timeout(GROQ_CONFIG.timeout_seconds, connect=20.0),
    )

    client = OpenAI(
        api_key=GROQ_CONFIG.api_key,
        base_url=GROQ_CONFIG.base_url,
        http_client=http_client,
        max_retries=2,
    )

    try:
        response = client.chat.completions.create(
            model=GROQ_CONFIG.model_name,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            timeout=GROQ_CONFIG.timeout_seconds,
        )
    except APIConnectionError as exc:
        # Espone la causa reale (es. errore DNS, TLS, timeout TCP) invece
        # del generico "Connection error" che Streamlit mostrerebbe altrimenti.
        causa = repr(exc.__cause__) if exc.__cause__ else str(exc)
        raise GroqResponseError(
            "Impossibile raggiungere il server di Groq (api.groq.com) dalla "
            f"rete di hosting attuale. Dettaglio tecnico: {causa}. "
            "Verifica che l'host consenta connessioni HTTPS in uscita verso "
            "api.groq.com sulla porta 443."
        ) from exc

    return response.choices[0].message.content or ""


# ---------------------------------------------------------------------------
# PARSING / VALIDAZIONE RISPOSTA
# ---------------------------------------------------------------------------
def _strip_markdown_fences(raw: str) -> str:
    """Rimuove eventuali blocchi ```json ... ``` che il modello potrebbe
    aggiungere nonostante le istruzioni."""
    match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", raw, re.DOTALL)
    if match:
        return match.group(1)
    return raw.strip()


def _parse_voci_response(raw: str) -> Dict:
    """Valida la risposta grezza dell'AI (estrazione voci) e restituisce
    un dict con 'voci' (list), 'gestione_fiscale' (dict) e 'note' (list)."""
    cleaned = _strip_markdown_fences(raw)

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise GroqResponseError(
            "La risposta ricevuta da Groq non e' un JSON valido. "
            "Riprova; se il problema persiste, verifica il prompt o il modello configurato."
        ) from exc

    if "voci" not in data:
        raise GroqResponseError(
            "La risposta di Groq non contiene la sezione attesa: 'voci'."
        )

    voci_raw = data.get("voci", [])
    if not isinstance(voci_raw, list):
        raise GroqResponseError("Formato risposta non valido: 'voci' deve essere una lista.")

    voci: List[Dict] = []
    for v in voci_raw:
        if not isinstance(v, dict):
            continue
        try:
            importo = float(v.get("importo", 0.0))
        except (TypeError, ValueError):
            importo = 0.0
        voci.append({
            "descrizione": str(v.get("descrizione", "")).strip(),
            "codice_conto": str(v.get("codice_conto", "") or "").strip(),
            "importo": importo,
            "tipo": str(v.get("tipo", "")).strip().lower(),
        })

    gestione_fiscale_raw = data.get("gestione_fiscale", {})
    gestione_fiscale: Dict[str, float] = {}
    if isinstance(gestione_fiscale_raw, dict):
        for k, val in gestione_fiscale_raw.items():
            try:
                gestione_fiscale[k] = float(val)
            except (TypeError, ValueError):
                gestione_fiscale[k] = 0.0
    for categoria in GESTIONE_FISCALE_CATEGORIE:
        gestione_fiscale.setdefault(categoria, 0.0)

    return {
        "voci": voci,
        "gestione_fiscale": gestione_fiscale,
        "note": list(data.get("note", []) or []),
    }


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA
# ---------------------------------------------------------------------------
def riclassifica_bilancio(anno: str, testi_pdf: List[str]) -> RiclassificazioneResult:
    """
    Punto di ingresso principale del modulo: prende il testo estratto dai
    PDF caricati, chiede a Groq di ESTRARRE (non classificare) ogni voce, e
    poi usa il classificatore deterministico (classificatore.py) per
    assegnare ciascuna voce alla categoria corretta.

    Solleva GroqConfigError se la chiave API non e' impostata, e
    GroqResponseError se la risposta non rispetta il formato atteso.
    """
    # Import locale per evitare un ciclo di import (classificatore.py
    # importa RiclassificazioneResult/Voce da questo stesso modulo).
    from classificatore import classifica_voci, filtra_subtotali_gerarchia

    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(anno, testi_pdf)

    raw_response = _call_model(system_prompt, user_prompt)
    dati = _parse_voci_response(raw_response)

    # Deduplicazione gerarchica DETERMINISTICA: l'AI ha estratto anche le
    # righe di subtotale (es. "3.66" oltre a "3.66.05.001"/"3.66.05.002"),
    # di proposito - non le ha gia' filtrate lei stessa perche' quel
    # giudizio si e' rivelato inaffidabile. Qui, confrontando i codici
    # conto, si scartano matematicamente le righe che sono subtotali di
    # altre righe piu' specifiche presenti nell'estrazione.
    voci_filtrate = filtra_subtotali_gerarchia(dati["voci"])

    risultato = classifica_voci(voci_filtrate)
    risultato.note = list(risultato.note) + list(dati["note"])

    # La gestione fiscale non e' ricavata da singole voci del bilancio, ma
    # estratta direttamente dall'AI come valore aggregato per categoria
    # (es. IVA vendite VE26): viene "impacchettata" come una singola Voce
    # sintetica per categoria, per restare coerente con la struttura
    # Dict[str, List[Voce]] usata anche da entrate/uscite.
    risultato.gestione_fiscale = {
        categoria: [Voce(descrizione="(valore aggregato dal bilancio)", importo=valore)]
        for categoria, valore in dati["gestione_fiscale"].items()
    }

    return risultato


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA: SOLA ESTRAZIONE (per il flusso di revisione manuale)
# ---------------------------------------------------------------------------
def estrai_voci_bilancio(anno: str, testi_pdf: List[str]) -> Dict:
    """
    Come riclassifica_bilancio, ma SENZA applicare la classificazione
    automatica per categoria: restituisce le voci grezze (gia' passate
    dalla deduplicazione gerarchica) cosi' come sono, per permettere
    all'utente di confermarle/correggerle una per una nell'interfaccia
    (flusso a popup).

    Restituisce un dict con:
      - "voci": lista di dict {"descrizione", "codice_conto", "importo", "tipo"}
      - "gestione_fiscale": dict {categoria: valore_numerico}
      - "note": lista di osservazioni testuali
    """
    from classificatore import filtra_subtotali_gerarchia

    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(anno, testi_pdf)

    raw_response = _call_model(system_prompt, user_prompt)
    dati = _parse_voci_response(raw_response)

    dati["voci"] = filtra_subtotali_gerarchia(dati["voci"])

    return dati
