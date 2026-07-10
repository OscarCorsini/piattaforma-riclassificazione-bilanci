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

Le regole sono state costruite analizzando OLTRE 100 bilanci reali di
aziende agricole italiane (bilanci di verifica, conti economici, situazioni
contabili di aziende zootecniche, vitivinicole, cerealicole) presenti nella
cartella condivisa dall'utente, per intercettare la terminologia realmente
usata (es. "COSTI PER FARMACI VETERINARI", "LATTE CRUDO ALLA STALLA",
"COSTI PER OLI E LUBRIFICANTI", "FORMAGGIO DA LATTE CAPRINO") e non solo
quella di un singolo documento di esempio.
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
# VOCI DA SCARTARE: righe di riepilogo/risultato che l'AI a volte estrae per
# errore nonostante le istruzioni (es. "UTILE D'ESERCIZIO", "TOTALE RICAVI",
# "REDDITO IMPONIBILE"), o voci di investimento/capex (es. "BENI
# STRUMENTALI") che non sono costi/ricavi operativi e non vanno sommate nel
# conto economico riclassificato. Se una voce corrisponde a uno di questi
# pattern, viene ignorata del tutto (non finisce ne' in una categoria
# specifica ne' nel generico "Altro"/"Oneri diversi").
# ---------------------------------------------------------------------------
PATTERN_DA_SCARTARE = re.compile(
    r"^(totale|subtotale|progressivo|a pareggio|utile d.?esercizio|"
    r"perdita d.?esercizio|reddito imponibile|beni strumentali|"
    r"totale attivita|totale passivita|totale generale|totale costi|"
    r"totale ricavi|risultato (d.?esercizio|ante imposte|operativo))\b"
)


# ---------------------------------------------------------------------------
# DEDUPLICAZIONE GERARCHICA DEI CODICI CONTO (deterministica)
# ---------------------------------------------------------------------------
# Molti bilanci di verifica hanno una struttura a piu' livelli dove un
# codice piu' corto e' il SUBTOTALE dei codici piu' lunghi che iniziano
# allo stesso modo (es. "3.66" e' il subtotale di "3.66.05.001" e
# "3.66.05.002"). In precedenza si chiedeva all'AI di riconoscere DA SOLA
# questi subtotali e scartarli durante l'estrazione: si e' rivelato
# inaffidabile (il modello a volte estraeva comunque la riga aggregata
# invece del dettaglio, con l'effetto pratico che gli importi finivano
# nelle categorie generiche invece che in quelle specifiche).
#
# Ora l'AI estrae DELIBERATAMENTE tutte le righe, comprese quelle
# aggregate, insieme al loro codice conto. Questa funzione elimina poi in
# modo puramente matematico (confronto di prefissi tra codici) le righe
# che risultano essere il subtotale di altre righe piu' specifiche
# presenti nella stessa estrazione, tenendo solo le righe "foglia"
# (quelle senza ulteriori sotto-codici tra le voci estratte).
def _tokenizza_codice(codice: str) -> tuple:
    """Scompone un codice conto (es. '3.66.05.001' o '66/25/010') nei suoi
    segmenti, per poter confrontare prefissi indipendentemente dal
    separatore usato nel documento (punto, slash, trattino)."""
    parti = re.split(r"[.\-/\s]+", str(codice).strip())
    return tuple(p for p in parti if p)


def filtra_subtotali_gerarchia(voci: List[Dict]) -> List[Dict]:
    """
    Rimuove dalla lista di voci quelle il cui codice_conto e' un prefisso
    del codice_conto di un'ALTRA voce nella stessa lista (cioe' sono il
    subtotale di quell'altra voce, piu' specifica). Le voci senza
    codice_conto (o con un codice che non compare come prefisso di
    nessun'altra) vengono mantenute cosi' come sono, perche' per loro non
    e' possibile determinare la gerarchia in modo affidabile.
    """
    codici_tokenizzati = []
    for v in voci:
        codice = str(v.get("codice_conto", "") or "").strip()
        if codice:
            codici_tokenizzati.append(_tokenizza_codice(codice))

    def ha_figli(token: tuple) -> bool:
        for altro in codici_tokenizzati:
            if altro != token and len(altro) > len(token) and altro[: len(token)] == token:
                return True
        return False

    risultato = []
    for v in voci:
        codice = str(v.get("codice_conto", "") or "").strip()
        if not codice:
            risultato.append(v)
            continue
        token = _tokenizza_codice(codice)
        if ha_figli(token):
            # E' un subtotale di righe piu' specifiche presenti
            # nell'estrazione: viene scartato per evitare il doppio
            # conteggio dello stesso importo.
            continue
        risultato.append(v)

    return risultato


# ---------------------------------------------------------------------------
# REGOLE DI CLASSIFICAZIONE USCITE (ordine = priorita': la prima regola che
# trova corrispondenza vince; le regole piu' specifiche vanno per prime,
# quelle generiche/di ripiego per ultime).
# ---------------------------------------------------------------------------
REGOLE_USCITE: List[Tuple[str, str]] = [
    (r"carburant|gasoli|benzin|\bdiesel\b|lubrificant|combustibil",
     "Carburanti"),
    (r"mangim|forag|\bfien\b|insilat|zootecni.*aliment|integrator|"
     r"\bmais\b|crusca|farina.*zootecnic|\bpaglia\b|erba medica|trinciat",
     "Mangimi e Foraggi"),
    (r"sement|semin|floro.?viv|vivais|barbatell",
     "Sementi"),
    (r"energia elettric|\benel\b|\beni\b|\ba2a\b|\bedison\b|sorgenia|"
     r"bolletta elettric|gas riscald|riscaldamento|acqua.*zootecnic|"
     r"acqua uso", "Energia elettrica"),
    (r"manut|riparaz|ricambi|officina|pneumatic|gomme e|ghiaia|pietre",
     "Manutenzioni"),
    (r"lavoraz.*terzi|conto\s*terzi|contoterzis.*(passiv|terzi)|mietitrebb",
     "Lavorazioni c/terzi"),
    (r"\bleasing\b", "Leasing"),
    (r"affitt|canone.*(terreno|fabbricat|immobil)", "Affitti"),
    (r"assicura|polizza|grandine", "Assicurazioni"),
    (r"consorzi.*bonifica|canone.*irrigu|acqua irrigu|taglia.*acqua",
     "Taglie acqua irrigua"),
    (r"salari|stipend|dipendent|personale", "Salari lordi dip."),
    (r"prelievi titolare|compenso.*titolare|compenso.*soci",
     "Prelievi titolare"),
    (r"contributi.*(inps|prev|scau|sindacal)|cassa previdenza",
     "Contributi prev."),
    (r"consulenz|assistenz|professional|servizi|utenz|trasport|"
     r"noleggio|comp\.?\s*prof|telefon|contabilita|tenuta paghe|"
     r"commercialist|pubblicit|fiere e mercati|spedizion|analisi.*labor|"
     r"laboratori|smaltimento|formazione|certificazion|controllo.*cee|"
     r"spese amministrative|disinfestazion|profilassi|pulizi",
     "Servizi"),
    (r"sanif|indument|dispositiv.*protezion|altri acquist|non inerent|"
     r"altri prod|divise|vestiario|detersiv|detergent|cancelleria|"
     r"imballagg|material.*consumo|scatole|borse|vasi e capsule|"
     r"tappi e capsule|spese accessorie|\baltri\b",
     "Altri"),
    (r"merci|materie prim|prodott|concim|fitofarmac|acquist|animali|"
     r"antiparassit|farmaci veterinari|\bveterinari\b|disinfettant|"
     r"seme animale|fecondazione|trucioli",
     "Materie Prime e Merci"),
]

REGOLE_ENTRATE: List[Tuple[str, str]] = [
    (r"vendita|vendite|corrispettiv|cession|merci conto vendite|"
     r"bottiglie conto vendite|\blatte\b|formaggio|\bcarne\b|salame|"
     r"prosciutto|conserve|yogurt|olio di oliva|\bvino\b|vitell|"
     r"\bbovin|\bsuin|ovini|caprini|ingrasso|confezione di formaggi|"
     r"conferimento", "Corrispettivi normali"),
    (r"agrituris|contoterzis.*attiv|vendita energia|fotovoltaic|prestazio|"
     r"noleggi|riaddebito|fitti attivi|deposito|\bletame\b",
     "Attivita' connessa"),
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
    con 0.0 se non ha ricevuto alcuna voce. Le righe di riepilogo/risultato
    (es. "TOTALE RICAVI", "UTILE D'ESERCIZIO") vengono riconosciute e
    scartate, per evitare di sommarle per errore nel conto riclassificato.
    """
    entrate: Dict[str, float] = {c: 0.0 for c in ENTRATE_CATEGORIE}
    uscite: Dict[str, float] = {c: 0.0 for c in USCITE_CATEGORIE}
    # Traccia anche, per ciascuna categoria, l'elenco delle singole voci
    # originali che vi sono state assegnate (descrizione + importo): non
    # serve per il calcolo del totale (gia' fatto sopra), ma permette
    # all'interfaccia di mostrare un dettaglio verificabile di "quale voce
    # e' finita in quale categoria", utile per controllare rapidamente che
    # la classificazione sia corretta su un bilancio reale.
    dettaglio_entrate: Dict[str, List[Dict]] = {c: [] for c in ENTRATE_CATEGORIE}
    dettaglio_uscite: Dict[str, List[Dict]] = {c: [] for c in USCITE_CATEGORIE}
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

        if PATTERN_DA_SCARTARE.search(_normalizza(descrizione)):
            continue

        tipo_normalizzato = "entrata" if tipo.startswith("entrat") else "uscita"
        categoria = _classifica_una_voce(descrizione, tipo_normalizzato)

        if tipo_normalizzato == "entrata":
            entrate[categoria] = entrate.get(categoria, 0.0) + importo
            dettaglio_entrate.setdefault(categoria, []).append(
                {"descrizione": descrizione, "importo": importo}
            )
        else:
            uscite[categoria] = uscite.get(categoria, 0.0) + importo
            dettaglio_uscite.setdefault(categoria, []).append(
                {"descrizione": descrizione, "importo": importo}
            )

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
        dettaglio_entrate=dettaglio_entrate,
        dettaglio_uscite=dettaglio_uscite,
    )
