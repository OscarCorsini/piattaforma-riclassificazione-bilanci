# -*- coding: utf-8 -*-
"""
config.py
=========
Configurazione centrale della piattaforma di riclassificazione finanziaria.

Qui sono definiti:
  - Le macro-categorie contabili (Entrate / Uscite / Gestione Fiscale)
  - La mappatura "categoria -> cella Excel" per ciascun anno gestito dal template
  - Le impostazioni di connessione a Groq (motore AI)

IMPORTANTE:
  Il mapping "CELL_MAPPING" sotto riportato e' un ESEMPIO basato sulla struttura
  descritta nei requisiti. Deve essere adattato alle celle REALI del file
  "Template.xlsx" fornito dall'azienda (foglio e indirizzo cella).
  Il modo piu' semplice per adattarlo e' aprire il template, individuare la riga
  di ciascuna voce e la colonna corrispondente all'anno, e aggiornare i valori
  qui sotto (es. "C15" per anno 2024, "D15" per anno 2025, ecc.).
"""

from dataclasses import dataclass, field
from typing import Dict, List


# ---------------------------------------------------------------------------
# 1. MACRO-CATEGORIE DI RICLASSIFICAZIONE
# ---------------------------------------------------------------------------
ENTRATE_CATEGORIE: List[str] = [
    "Corrispettivi normali",
    "Attivita' connessa",
    "PAC e contributi pubblici",
    "Gestione fiscale",
    "Variazione rimanenze",
    "Altro",
]

USCITE_CATEGORIE: List[str] = [
    "Materie prime e merci",
    "Servizi",
    "Leasing",
    "Altri acquisti",
    "Affitti",
    "Salari lordi dip.",
    "Prelievi titolare",
    "Contributi prev.",
    "Assicurazioni",
    "Taglie acqua irrigua",
    "Oneri diversi di gestione",
]

GESTIONE_FISCALE_CATEGORIE: List[str] = [
    "Iva vendite (VE26)",
    "Iva acquisti (VF27)",
    "Imposta dovuta (VL3)",
    "Netto gestione",
]


# ---------------------------------------------------------------------------
# 2. MAPPING CELLE EXCEL (foglio + cella per ciascun anno)
# ---------------------------------------------------------------------------
SHEET_NAME_RICLASSIFICAZIONE = "Riclassificazione"
SHEET_NAME_GESTIONE_FISCALE = "Riclassificazione"  # stessa scheda, tabella laterale

# Mappa: {anno: {categoria: cella}}
# ATTENZIONE: valori di esempio, da adattare al template reale.
CELL_MAPPING: Dict[str, Dict[str, str]] = {
    "2024": {
        "Corrispettivi normali": "C6",
        "Attivita' connessa": "C7",
        "PAC e contributi pubblici": "C8",
        "Gestione fiscale": "C9",
        "Variazione rimanenze": "C10",
        "Altro": "C11",
        "Materie prime e merci": "C15",
        "Servizi": "C16",
        "Leasing": "C17",
        "Altri acquisti": "C18",
        "Affitti": "C19",
        "Salari lordi dip.": "C20",
        "Prelievi titolare": "C21",
        "Contributi prev.": "C22",
        "Assicurazioni": "C23",
        "Taglie acqua irrigua": "C24",
        "Oneri diversi di gestione": "C25",
        "Iva vendite (VE26)": "H6",
        "Iva acquisti (VF27)": "H7",
        "Imposta dovuta (VL3)": "H8",
        "Netto gestione": "H9",
    },
    "2025": {
        "Corrispettivi normali": "D6",
        "Attivita' connessa": "D7",
        "PAC e contributi pubblici": "D8",
        "Gestione fiscale": "D9",
        "Variazione rimanenze": "D10",
        "Altro": "D11",
        "Materie prime e merci": "D15",
        "Servizi": "D16",
        "Leasing": "D17",
        "Altri acquisti": "D18",
        "Affitti": "D19",
        "Salari lordi dip.": "D20",
        "Prelievi titolare": "D21",
        "Contributi prev.": "D22",
        "Assicurazioni": "D23",
        "Taglie acqua irrigua": "D24",
        "Oneri diversi di gestione": "D25",
        "Iva vendite (VE26)": "I6",
        "Iva acquisti (VF27)": "I7",
        "Imposta dovuta (VL3)": "I8",
        "Netto gestione": "I9",
    },
}

# Anni selezionabili in interfaccia.
ANNI_DISPONIBILI: List[str] = list(CELL_MAPPING.keys())


# ---------------------------------------------------------------------------
# 3. CONFIGURAZIONE GROQ
# ---------------------------------------------------------------------------
@dataclass
class GroqConfig:
    """
    Parametri di connessione all'API di Groq (motore AI, gratuito, senza
    carta di credito richiesta).

    >>> INSERIRE QUI LA CHIAVE API <<<
    Genera una chiave gratuita su https://console.groq.com/keys con un
    account email o Google, poi impostala come GROQ_API_KEY (variabile
    d'ambiente in locale, oppure nei Secrets di Streamlit Community Cloud).
    Non scrivere mai la chiave direttamente nel codice.
    """

    api_key: str = field(default_factory=lambda: _env("GROQ_API_KEY", ""))
    model_name: str = field(default_factory=lambda: _env("GROQ_MODEL", "llama-3.3-70b-versatile"))
    base_url: str = "https://api.groq.com/openai/v1"
    timeout_seconds: int = 90


def _env(key: str, default: str) -> str:
    # Legge un parametro di configurazione da (in ordine di priorita'):
    #   1. st.secrets - usato da Streamlit Community Cloud, dove le
    #      credenziali si inseriscono nella dashboard (Settings > Secrets)
    #      invece che come variabili d'ambiente del sistema operativo.
    #   2. Variabili d'ambiente del sistema - usato per l'esecuzione locale
    #      o su altri hosting.
    #   3. Valore di default.
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    import os
    return os.environ.get(key, default)


# Istanza di configurazione usata dall'applicazione.
GROQ_CONFIG = GroqConfig()
