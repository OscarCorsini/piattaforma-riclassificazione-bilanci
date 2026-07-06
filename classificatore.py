# -*- coding: utf-8 -*-
"""
classificatore.py
==================
Classificazione DETERMINISTICA (basata su parole chiave, non sull'AI) delle
singole voci di un bilancio di verifica nelle macro-categorie del template.

Perche' questo modulo esiste:
-----------------------------
Chiedere a un modello linguistico di leggere un intero bilancio di verifica
E contemporaneamente sommare/classificare decine di voci in 20+ categorie in
un solo colpo si e' rivelato inaffidabile: il modello tende ad accorpare le
voci sotto poche categorie generiche ("Servizi", "Materie Prime e Merci"),
anche quando il documento originale elenca chiaramente le singole voci
(es. "ENERGIA ELETTRICA", "MERCI C/ACQUISTI - CARBURANTI") con un proprio
codice conto e una propria descrizione.

La soluzione adottata e' un approccio IBRIDO:
  1. L'AI (groq_client.py) si limita a ESTRARRE ogni singola voce del
     bilancio come testo (descrizione + importo + entrata/uscita) - un
     compito semplice e meccanico che i modelli linguistici svolgono in
     modo molto piu' affidabile rispetto alla classificazione aggregata.
  2. Questo modulo assegna poi CIASCUNA voce estratta alla categoria
     corretta confrontando la descrizione con un elenco di parole chiave,
     in ordine di priorita' (dalla piu' specifica alla piu' generica).
     Essendo puro codice Python, il risultato e' sempre deterministico e
     ripetibile: stesso testo in ingresso, stessa classificazione in uscita.

Le regole sono state costruite e verificate sulla struttura tipica di un
bilancio di verifica italiano (codici conto tipo "66/25/010" seguiti da una
descrizione abbreviata in maiuscolo, es. "MERCI C/ACQUISTI - CARBURANTI",
"ACQ. CEREALI E FORAGGI", "LAVORAZ.DI TERZI P/PROD.SERVIZI").
"""

import re
import unicodedata
from typing import Dict, List, Tuple

from groq_client import RiclassificazioneResult
from config import ENTRATE_CATEGORIE, USCITE_CATEGORIE


def _normalizza(testo) -> str:
    """Minuscolo, senza accenti: rende il pattern-matching robusto a
    maiuscole/minuscole e accenti (es. 'attivita' vs 'attivita'')."""
    if testo is None:
        return ""
    t = unicodedata.normalize("NFKD", str(testo))
    t = t.encode("ascii", "ignore").decode("ascii")
    return t.lower()


# ---------------------------------------------------------------------------
# REGOLE DI CLASSIFICAZIONE USCITE (ordine = priorita': la prima regola che
# trova corrispondenza vince; le regole piu' specifiche vanno per prime,
# quelle generiche/di ripiego per ultime).
# ---------------------------------------------------------------------------
REGOLE_USCITE: List[Tuple[str, str]] = [
    (r"carburant|gasoli|benzin|\bdiesel\b", "Carburanti"),
    (r"mangim|forag|\bfien\b|insilat|zootecni.*aliment", "Mangimi e Foraggi"),
    (r"sement|semin|floro.?viv|vivais", "Sementi"),
    (r"energia elettric|\benel\b|\beni\b|\ba2a\b|\bedison\b|sorgenia|"
     r"bolletta elettric|gas riscald|riscaldamento", "Energia elettrica"),
    (r"manut|riparaz|ricambi|officina", "Manutenzioni"),
    (r"lavoraz.*terzi|conto\s*terzi|contoterzis.*(passiv|terzi)|mietitrebb", "Lavorazioni c/terzi"),
    (r"\bleasing\b", "Leasing"),
    (r"affitt|canone.*(terreno|fabbricat|immobil)", "Affitti"),
    (r"assicura|polizza|grandine", "Assicurazioni"),
    (r"consorzi.*bonifica|canone.*irrigu|acqua irrigu|taglia.*acqua", "Taglie acqua irrigua"),
    (r"salari|stipend|dipendent|personale", "Salari lordi dip."),
    (r"prelievi titolare|compenso.*titolare|compenso.*soci", "Prelievi titolare"),
    (r"contributi.*(inps|prev|scau|sindacal)", "Contributi prev."),
    (r"consulenz|assistenz|professional|servizi|utenz|trasport|noleggio|comp\.?\s*prof",
     "Servizi"),
    (r"sanif|indument|dispositiv.*protezion|altri acquist|non inerent|altri prod|\baltri\b",
     "Altri"),
    (r"merci|materie prim|prodott|concim|fitofarmac|acquist|animali",
     "Materie Prime e Merci"),
]

REGOLE_ENTRATE: List[Tuple[str, str]] = [
    (r"vendita|vendite|corrispettiv|cession|merci conto vendite|bottiglie conto vendite",
     "Corrispettivi normali"),
    (r"agrituris|contoterzis.*attiv|vendita energia|fotovoltaic|prestazio|noleggi|"
     r"riaddebito|fitti attivi|deposito", "Attivita' connessa"),
    (r"\bpac\b|\bpsr\b|contribut", "PAC e contributi pubblici"),
    (r"\biva\b|imposta|fiscal", "Gestione fiscale"),
]


def _classifica_una_voce(descrizione: str, tipo: str) -> str:
    """Restituisce la categoria (stringa esatta come in config.py) piu'
    appropriata per una singola voce, in base a parole chiave nella sua
    descrizione. Se nessuna regola specifica corrisponde, restituisce la
    categoria di ripiego generica."""
    testo = _normalizza(descrizione)
    regole = REGOLE_ENTRATE if tipo == "entrata" else REGOLE_USCITE

    for pattern, categoria in regole:
        if re.search(pattern, testo):
            return categoria

    return "Altro" if tipo == "entrata" else "Oneri diversi di gestione"


def classifica_voci(voci: List[Dict]) -> RiclassificazioneResult:
    """
    Prende la lista di voci grezze estratte dall'AI (ognuna con
    'descrizione', 'importo', 'tipo': 'entrata'|'uscita') e restituisce un
    RiclassificazioneResult con gli importi gia' sommati per categoria,
    usando la classificazione deterministica basata su parole chiave.

    Ogni categoria configurata (config.py) e' sempre presente nel risultato,
    con 0.0 se non ha ricevuto alcuna voce.
    """
    entrate: Dict[str, float] = {c: 0.0 for c in ENTRATE_CATEGORIE}
    uscite: Dict[str, float] = {c: 0.0 for c in USCITE_CATEGORIE}
    note: List[str] = []

    for voce in voci:
        descrizione = voce.get("descrizione", "")
        tipo = str(voce.get("tipo", "")).strip().lower()
        try:
            importo = float(voce.get("importo", 0.0))
        except (TypeError, ValueError):
            importo = 0.0

        if not descrizione or importo == 0.0:
            continue

        tipo_normalizzato = "entrata" if tipo.startswith("entrat") else "uscita"
        categoria = _classifica_una_voce(descrizione, tipo_normalizzato)

        if tipo_normalizzato == "entrata":
            entrate[categoria] = entrate.get(categoria, 0.0) + importo
        else:
            uscite[categoria] = uscite.get(categoria, 0.0) + importo

        catchall = "Altro" if tipo_normalizzato == "entrata" else "Oneri diversi di gestione"
        if categoria == catchall:
            note.append(
                f"Voce '{descrizione}' non riconosciuta da nessuna regola "
                f"specifica: inserita in '{categoria}'."
            )

    return RiclassificazioneResult(
        entrate=entrate,
        uscite=uscite,
        gestione_fiscale={},
        note=note,
    )
