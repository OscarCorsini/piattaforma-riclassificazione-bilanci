# -*- coding: utf-8 -*-
"""
copilot_client.py
==================
Modulo di integrazione con Microsoft Copilot (motore AI aziendale).

Questo modulo:
  1. Costruisce un prompt strutturato con la logica di riclassificazione
     contabile agricola richiesta.
  2. Invia il testo estratto dai PDF a Copilot chiedendo una risposta in
     formato JSON rigido.
  3. Valida e normalizza la risposta prima di restituirla al chiamante.

>>> DOVE INSERIRE LE CREDENZIALI AZIENDALI <<<
Le credenziali (endpoint, api_key, deployment) sono lette da `config.py`
(classe CopilotConfig), che a sua volta le legge da variabili d'ambiente:
    COPILOT_ENDPOINT, COPILOT_API_KEY, COPILOT_DEPLOYMENT, COPILOT_API_VERSION

Impostale prima di avviare l'app, ad esempio in PowerShell:
    $env:COPILOT_ENDPOINT   = "https://<tuo-endpoint>.openai.azure.com/"
    $env:COPILOT_API_KEY    = "<la-tua-chiave>"
    $env:COPILOT_DEPLOYMENT = "<nome-deployment>"

oppure creando un file ".env" nella cartella del progetto (richiede
`python-dotenv`, non incluso di default: se lo usi, aggiungi in cima ad
app.py le righe:
    from dotenv import load_dotenv
    load_dotenv()

NOTA SULL'SDK:
Questo modulo usa la libreria `openai` (>=1.0), compatibile con Azure OpenAI
Service, che è il motore su cui si basa comunemente l'integrazione
aziendale "Microsoft Copilot" / "Copilot Studio" via Azure AI Foundry.
Se la tua organizzazione espone Copilot tramite un connettore differente
(es. Microsoft Graph Copilot API, o un gateway REST interno), sostituisci
il corpo della funzione `_call_model()` più sotto con la chiamata
equivalente, mantenendo invariata la firma della funzione e il formato di
ritorno (stringa JSON).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Dict, List, Optional

from config import (
    COPILOT_CONFIG,
    ENTRATE_CATEGORIE,
    USCITE_CATEGORIE,
    GESTIONE_FISCALE_CATEGORIE,
)


class CopilotConfigError(Exception):
    """Credenziali/endpoint mancanti o non validi."""


class CopilotResponseError(Exception):
    """La risposta di Copilot non è nel formato JSON atteso."""


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

    return f"""Sei un assistente esperto di contabilità agraria e bilanci d'esercizio.
Il tuo compito è leggere una situazione contabile (bilancio di verifica,
mastrini, estratto conto economico) fornita come testo estratto da PDF e
riclassificare ogni voce esattamente nelle seguenti macro-categorie,
sommando gli importi quando più voci originali appartengono alla stessa
macro-categoria.

ENTRATE:
{entrate}

USCITE:
{uscite}

GESTIONE FISCALE (dati IVA/imposte, se presenti nel documento):
{fiscale}

REGOLE:
1. Analizza tutte le voci di ricavo e costo presenti nel testo.
2. Assegna ciascuna voce alla macro-categoria più appropriata secondo la
   logica del settore agricolo (es. vendita latte/carne/cereali ->
   "Corrispettivi normali"; agriturismo/contoterzismo attivo/vendita
   energia -> "Attività connessa"; PSR/PAC/contributi regionali -> "PAC e
   contributi pubblici"; canoni leasing macchinari -> "Leasing"; affitto
   terreni/fabbricati -> "Affitti"; costo dipendenti -> "Salari lordi
   dip."; compensi titolare/soci -> "Prelievi titolare"; INPS/Ex-Scau ->
   "Contributi prev."; polizze grandine/RC -> "Assicurazioni"; consorzi di
   bonifica -> "Taglie acqua irrigua"; tutto il resto non classificabile
   -> "Altro" o "Oneri diversi di gestione").
3. Se una voce è ambigua o non chiaramente riconducibile a una categoria,
   inseriscila in "Altro" (per le entrate) o "Oneri diversi di gestione"
   (per le uscite), e segnalalo in "note".
4. Se non trovi dati per una categoria, restituisci 0.0 per quella
   categoria (non ometterla).
5. Tutti gli importi devono essere numeri (float), positivi, espressi in
   euro, senza simboli di valuta né separatori delle migliaia.
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

Testo estratto dai documenti contabili (uno o più PDF concatenati):

{testo_concatenato}

Analizza il contenuto sopra e restituisci il JSON di riclassificazione
richiesto secondo le istruzioni di sistema."""


# ---------------------------------------------------------------------------
# CHIAMATA AL MODELLO
# ---------------------------------------------------------------------------
def _call_model(system_prompt: str, user_prompt: str) -> str:
    """
    Effettua la chiamata al modello Copilot/Azure OpenAI e restituisce il
    contenuto testuale grezzo della risposta.

    >>> SOSTITUIRE QUESTO BLOCCO CON LA CHIAMATA AL CONNETTORE AZIENDALE <<<
    se non si utilizza Azure OpenAI SDK.
    """
    if not COPILOT_CONFIG.endpoint or not COPILOT_CONFIG.api_key:
        raise CopilotConfigError(
            "Credenziali Copilot non configurate. Imposta le variabili d'ambiente "
            "COPILOT_ENDPOINT e COPILOT_API_KEY (vedi commenti in config.py / "
            "copilot_client.py) prima di avviare l'elaborazione."
        )

    try:
        # Import locale per non rendere obbligatoria la dipendenza se si
        # sostituisce con un altro SDK/connettore.
        from openai import AzureOpenAI
    except ImportError as exc:
        raise CopilotConfigError(
            "Libreria 'openai' non installata. Esegui: pip install openai"
        ) from exc

    client = AzureOpenAI(
        azure_endpoint=COPILOT_CONFIG.endpoint,
        api_key=COPILOT_CONFIG.api_key,
        api_version=COPILOT_CONFIG.api_version,
    )

    response = client.chat.completions.create(
        model=COPILOT_CONFIG.deployment_name,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        timeout=COPILOT_CONFIG.timeout_seconds,
    )

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


def _parse_response(raw: str) -> RiclassificazioneResult:
    cleaned = _strip_markdown_fences(raw)

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise CopilotResponseError(
            "La risposta ricevuta da Copilot non è un JSON valido. "
            "Riprova; se il problema persiste, verifica il prompt o il modello configurato."
        ) from exc

    required_keys = {"entrate", "uscite", "gestione_fiscale"}
    if not required_keys.issubset(data.keys()):
        mancanti = required_keys - data.keys()
        raise CopilotResponseError(
            f"La risposta di Copilot non contiene le sezioni attese: {', '.join(mancanti)}."
        )

    def _to_float_dict(d) -> Dict[str, float]:
        result = {}
        if not isinstance(d, dict):
            raise CopilotResponseError("Formato risposta non valido: attesa una mappa categoria->valore.")
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

    Solleva CopilotConfigError se le credenziali non sono impostate, e
    CopilotResponseError se la risposta non rispetta il formato atteso.
    """
    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(anno, testi_pdf)

    raw_response = _call_model(system_prompt, user_prompt)
    return _parse_response(raw_response)
