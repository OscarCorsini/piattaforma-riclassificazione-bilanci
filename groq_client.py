# -*- coding: utf-8 -*-
"""
groq_client.py
==============
Modulo di integrazione con Groq (motore AI usato per la riclassificazione
contabile). Groq offre un'API compatibile con lo standard OpenAI, gratuita
e senza richiesta di carta di credito.

Questo modulo:
  1. Costruisce un prompt strutturato con la logica di riclassificazione
     contabile agricola richiesta.
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
from dataclasses import dataclass
from typing import Dict, List

from config import (
    GROQ_CONFIG,
    ENTRATE_CATEGORIE,
    USCITE_CATEGORIE,
    GESTIONE_FISCALE_CATEGORIE,
)


class GroqConfigError(Exception):
    """Chiave API mancante o non valida."""


class GroqResponseError(Exception):
    """La risposta di Groq non e' nel formato JSON atteso."""


@dataclass
class VoceDettaglio:
    descrizione: str
    importo: float


@dataclass
class RiclassificazioneResult:
    entrate: Dict[str, List[VoceDettaglio]]
    uscite: Dict[str, List[VoceDettaglio]]
    gestione_fiscale: Dict[str, List[VoceDettaglio]]
    note: List[str]

    def totale_entrate(self) -> Dict[str, float]:
        return {k: sum(v.importo for v in voci) for k, voci in self.entrate.items()}

    def totale_uscite(self) -> Dict[str, float]:
        return {k: sum(v.importo for v in voci) for k, voci in self.uscite.items()}

    def totale_gestione_fiscale(self) -> Dict[str, float]:
        return {k: sum(v.importo for v in voci) for k, voci in self.gestione_fiscale.items()}


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
riclassificare ogni voce nelle seguenti macro-categorie.

ENTRATE:
{entrate}

USCITE:
{uscite}

GESTIONE FISCALE (dati IVA/imposte, se presenti nel documento):
{fiscale}

REGOLE:
1. Analizza tutte le voci di ricavo e costo presenti nel testo.
2. Assegna ciascuna voce alla macro-categoria piu' appropriata secondo la
   logica del settore agricolo. Mantieni la descrizione originale della voce
   (la "microvoce") e il suo importo. Usa le categorie PIU' SPECIFICHE:
   - vendita latte/carne/cereali -> "Corrispettivi normali"
   - agriturismo/contoterzismo attivo/vendita energia -> "Attivita' connessa"
   - PSR/PAC/contributi regionali -> "PAC e contributi pubblici"
   - mangimi, foraggi, alimenti per il bestiame -> "Mangimi e Foraggi"
   - gasolio agricolo, benzina, carburanti per mezzi -> "Carburanti"
   - sementi, piantine, materiale di semina -> "Sementi"
   - concimi, fitofarmaci, materie prime generiche non altrimenti
     classificabili -> "Materie Prime e Merci"
   - altri acquisti minori non riconducibili alle voci sopra -> "Altri"
   - bollette elettriche, energia elettrica -> "Energia elettrica"
   - manutenzione macchinari, impianti, fabbricati -> "Manutenzioni"
   - lavorazioni conto terzi (contoterzismo passivo, es. mietitrebbiatura
     conto terzi) -> "Lavorazioni c/terzi"
   - consulenze, servizi professionali, altri servizi generici non
     riconducibili alle voci sopra -> "Servizi"
   - canoni leasing macchinari -> "Leasing"
   - affitto terreni/fabbricati -> "Affitti"
   - costo dipendenti -> "Salari lordi dip."
   - compensi titolare/soci -> "Prelievi titolare"
   - contributi INPS/Ex-Scau -> "Contributi prev."
   - polizze grandine/RC/assicurazioni -> "Assicurazioni"
   - consorzi di bonifica, canoni irrigui -> "Taglie acqua irrigua"
   - tutto il resto non classificabile -> "Altro" (entrate) o "Oneri
     diversi di gestione" (uscite)
3. Se una voce e' ambigua o non chiaramente riconducibile a una categoria,
   inseriscila in "Altro" (per le entrate) o "Oneri diversi di gestione"
   (per le uscite), e segnalalo in "note".
4. Se non trovi dati per una categoria, omettila dal JSON, non c'è bisogno di inserire array vuoti.
5. Tutti gli importi devono essere numeri (float), positivi, espressi in
   euro, senza simboli di valuta ne' separatori delle migliaia.
6. Estrai i dati di gestione fiscale SOLO se esplicitamente presenti nel
   documento.

FORMATO DI OUTPUT (OBBLIGATORIO):
Rispondi ESCLUSIVAMENTE con un oggetto JSON valido, senza testo aggiuntivo,
commenti, markdown o backtick, con questa struttura esatta:

{{
  "entrate": {{
    "<categoria>": [
      {{"descrizione": "<nome originale voce>", "importo": <valore_numerico>}}
    ]
  }},
  "uscite": {{
    "<categoria>": [
      {{"descrizione": "<nome originale voce>", "importo": <valore_numerico>}}
    ]
  }},
  "gestione_fiscale": {{
    "<categoria>": [
      {{"descrizione": "<nome originale voce>", "importo": <valore_numerico>}}
    ]
  }},
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
        from openai import OpenAI
    except ImportError as exc:
        raise GroqConfigError(
            "Libreria 'openai' non installata. Esegui: pip install openai"
        ) from exc

    client = OpenAI(
        api_key=GROQ_CONFIG.api_key,
        base_url=GROQ_CONFIG.base_url,
    )

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
        raise GroqResponseError(
            "La risposta ricevuta da Groq non e' un JSON valido. "
            "Riprova; se il problema persiste, verifica il prompt o il modello configurato."
        ) from exc

    required_keys = {"entrate", "uscite", "gestione_fiscale"}
    if not required_keys.issubset(data.keys()):
        mancanti = required_keys - data.keys()
        raise GroqResponseError(
            f"La risposta di Groq non contiene le sezioni attese: {', '.join(mancanti)}."
        )

    def _to_dettaglio_dict(d) -> Dict[str, List[VoceDettaglio]]:
        result = {}
        if not isinstance(d, dict):
            raise GroqResponseError("Formato risposta non valido: attesa una mappa categoria->lista dettagli.")
        for k, items in d.items():
            result[k] = []
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, dict):
                    continue
                desc = str(item.get("descrizione", "Voce non specificata"))
                try:
                    imp = float(item.get("importo", 0.0))
                except (TypeError, ValueError):
                    imp = 0.0
                result[k].append(VoceDettaglio(descrizione=desc, importo=imp))
        return result

    return RiclassificazioneResult(
        entrate=_to_dettaglio_dict(data.get("entrate", {})),
        uscite=_to_dettaglio_dict(data.get("uscite", {})),
        gestione_fiscale=_to_dettaglio_dict(data.get("gestione_fiscale", {})),
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

    Solleva GroqConfigError se la chiave API non e' impostata, e
    GroqResponseError se la risposta non rispetta il formato atteso.
    """
    system_prompt = _build_system_prompt()
    user_prompt = _build_user_prompt(anno, testi_pdf)

    raw_response = _call_model(system_prompt, user_prompt)
    return _parse_response(raw_response)
