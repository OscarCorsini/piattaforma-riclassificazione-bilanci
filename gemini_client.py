# -*- coding: utf-8 -*-
"""
gemini_client.py
================
Modulo di integrazione con Google Gemini (motore AI usato per la
riclassificazione contabile).

Questo modulo:
  1. Costruisce un prompt strutturato con la logica di riclassificazione
     contabile agricola richiesta.
  2. Invia il testo estratto dai PDF a Gemini chiedendo una risposta in
     formato JSON rigido.
  3. Valida e normalizza la risposta prima di restituirla al chiamante.

>>> DOVE INSERIRE LA CHIAVE API <<<
La chiave (GEMINI_API_KEY) e' letta da config.py (classe GeminiConfig),
che a sua volta la legge da st.secrets (Streamlit Community Cloud) o da
variabile d'ambiente (esecuzione locale).

Per generarla: vai su https://aistudio.google.com/apikey con il tuo
account Google, clicca "Create API key" e copiala. Il piano gratuito non
richiede una carta di credito.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Dict, List

from config import (
    GEMINI_CONFIG,
    ENTRATE_CATEGORIE,
    USCITE_CATEGORIE,
    GESTIONE_FISCALE_CATEGORIE,
)


class GeminiConfigError(Exception):
    """Chiave API mancante o non valida."""


class GeminiResponseError(Exception):
    """La risposta di Gemini non e' nel formato JSON atteso."""


@dataclass
class RiclassificazioneResult:
    entrate: Dict[str, float]
    uscite: Dict[str, float]
    gestione_fiscale: Dict[str, float]
    note: List[str]


# ---------------------------------------------------------------------------
# COSTRUZIONE DEL PROMPT
# ---------------------------------------------------------------------------
def _build_system_prompt() -> str:
    entrate = "\n".join(f"  - {c}" for c in ENTRATE_CATEGORIE)
    uscite = "\n".join(f"  - {c}" for c in USCITE_CATEGORIE)
    fiscale = "\n".join(f"  - {c}" for c in GESTIONE_FISCALE_CATEGORIE)

    return f"""Sei un assistente esperto di contabilita' agraria e bilanci d'esercizio.
Il tuo compito e' leggere una situazione contabile (bilancio di verifica,
mastrini, estratto conto economico) fornita come testo estratto da PDF e
riclassificare ogni voce esattamente nelle seguenti macro-categorie,
sommando gli importi quando piu' voci originali appartengono alla stessa
macro-categoria.

ENTRATE:
{entrate}

USCITE:
{uscite}

GESTIONE FISCALE (dati IVA/imposte, se presenti nel documento):
{fiscale}

REGOLE:
1. Analizza tutte le voci di ricavo e costo presenti nel testo.
2. Assegna ciascuna voce alla macro-categoria piu' appropriata secondo la
   logica del settore agricolo (es. vendita latte/carne/cereali ->
   "Corrispettivi normali"; agriturismo/contoterzismo attivo/vendita
   energia -> "Attivita' connessa"; PSR/PAC/contributi regionali -> "PAC e
   contributi pubblici"; canoni leasing macchinari -> "Leasing"; affitto
   terreni/fabbricati -> "Affitti"; costo dipendenti -> "Salari lordi
   dip."; compensi titolare/soci -> "Prelievi titolare"; INPS/Ex-Scau ->
   "Contributi prev."; polizze grandine/RC -> "Assicurazioni"; consorzi di
   bonifica -> "Taglie acqua irrigua"; tutto il resto non classificabile
   -> "Altro" o "Oneri diversi di gestione").
3. Se una voce e' ambigua o non chiaramente riconducibile a una categoria,
   inseriscila in "Altro" (per le entrate) o "Oneri diversi di gestione"
   (per le uscite), e segnalalo in "note".
4. Se non trovi dati per una categoria, restituisci 0.0 per quella
   categoria (non ometterla).
5. Tutti gli importi devono essere numeri (float), positivi, espressi in
   euro, senza simboli di valuta ne' separatori delle migliaia.
6. Estrai i dati di gestione fiscale (IVA vendite VE26, IVA acquisti VF27,
   imposta dovuta VL3, netto gestione) SOLO se esplicitamente presenti nel
   documento; altrimenti restituisci 0.0.

FORMATO DI OUTPUT (OBBLIGATORIO):
Rispondi ESCLUSIVAMENTE con un oggetto JSON valido, senza testo aggiuntivo,
commenti, markdown o backtick, con questa struttura esatta:

{{
  "entrate": {{ "<categoria>": <valore_numerico>, ... }},
  "uscite": {{ "<categoria>": <valore_numerico>, ... }},
  "gestione_fiscale": {{ "<categoria>": <valore_numerico>, ... }},
  "note": ["<eventuali osservazioni sulle voci ambigue o mancanti>"]
}}
"""


def _build_user_prompt(anno: str, documenti_testo: List[str]) -> str:
    testo_concatenato = "\n\n----- NUOVO DOCUMENTO -----\n\n".join(documenti_testo)
    return f"""Anno di riferimento: {anno}

Testo estratto dai documenti contabili (uno o piu' PDF concatenati):

{testo_concatenato}

Analizza il contenuto sopra e restituisci il JSON di riclassificazione
richiesto secondo le istruzioni di sistema."""


# ---------------------------------------------------------------------------
# CHIAMATA AL MODELLO
# ---------------------------------------------------------------------------
def _call_model(system_prompt: str, user_prompt: str) -> str:
    """
    Effettua la chiamata all'API di Google Gemini e restituisce il
    contenuto testuale grezzo della risposta.
    """
    if not GEMINI_CONFIG.api_key:
        raise GeminiConfigError(
            "Chiave API di Gemini non configurata. Genera una chiave gratuita su "
            "https://aistudio.google.com/apikey e impostala come GEMINI_API_KEY "
            "(vedi commenti in config.py / gemini_client.py) prima di avviare "
            "l'elaborazione."
        )

    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise GeminiConfigError(
            "Libreria 'google-generativeai' non installata. Esegui: "
            "pip install google-generativeai"
        ) from exc

    genai.configure(api_key=GEMINI_CONFIG.api_key)

    model = genai.GenerativeModel(
        model_name=GEMINI_CONFIG.model_name,
        system_instruction=system_prompt,
    )

    response = model.generate_content(
        user_prompt,
        generation_config=genai.types.GenerationConfig(
            temperature=0,
            response_mime_type="application/json",
        ),
        request_options={"timeout": GEMINI_CONFIG.timeout_seconds},
    )

    return response.text or ""


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


def _parse_response(raw: str) -> RiclassificazioneResult:
    cleaned = _strip_markdown_fences(raw)

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise GeminiResponseError(
            "La risposta ricevuta da Gemini non e' un JSON valido. "
            "Riprova; se il problema persiste, verifica il prompt o il modello configurato."
        ) from exc

    required_keys = {"entrate", "uscite", "gestione_fiscale"}
    if not required_keys.issubset(data.keys()):
        mancanti = required_keys - data.keys()
        raise GeminiResponseError(
            f"La risposta di Gemini non contiene le sezioni attese: {', '.join(mancanti)}."
        )

    def _to_float_dict(d) -> Dict[str, float]:
        result = {}
        if not isinstance(d, dict):
            raise GeminiResponseError("Formato risposta non valido: attesa una mappa categoria->valore.")
        for k, v in d.items():
            try:
                result[k] = float(v)
            except (TypeError, ValueError):
                result[k] = 0.0
        return result

    return RiclassificazioneResult(
        entrate=_to_float_dict(data.get("entrate", {})),
        uscite=_to_float_dict(data.get("uscite", {})),
        gestione_fiscale=_to_float_dict(data.get("gestione_fiscale", {})),
        note=list(data.get("note", []) or []),
    )


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA
# ---------------------------------------------------------------------------
def riclassifica_bilancio(anno: str, testi_pdf: List[str]) -> RiclassificazioneResult:
    """
    Punto di ingresso principale del modulo: prende il testo estratto dai
    PDF caricati e restituisce l'oggetto RiclassificazioneResult con i
    valori standardizzati pronti per essere scritti nel file Excel.

    Solleva GeminiConfigError se la chiave API non e' impostata, e
    GeminiResponseError se la risposta non rispetta il formato atteso.
    """
    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(anno, testi_pdf)

    raw_response = _call_model(system_prompt, user_prompt)
    return _parse_response(raw_response)
