# -*- coding: utf-8 -*-
"""
config.py
=========
Configurazione centrale della piattaforma di riclassificazione finanziaria.

Qui sono definiti:
  - Le macro-categorie contabili (Entrate / Uscite / Gestione Fiscale)
  - Il percorso del template Excel predefinito (fisso, incorporato nell'app)
  - Le impostazioni di connessione a Groq (motore AI)
  - La palette colori usata nell'interfaccia
"""

import os
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
    "Altro",
]

# Voci di uscita dettagliate: i nomi devono corrispondere ESATTAMENTE
# (a meno di maiuscole/minuscole) alle etichette presenti nel template
# Excel, perche' excel_writer.py cerca la riga giusta confrontando il
# testo normalizzato. Il template attuale ha alcune voci di "Acquisti
# ordinari" e "servizi" scomposte in dettaglio (es. Energia elettrica,
# Manutenzioni, Lavorazioni c/terzi, Mangimi e Foraggi, Carburanti,
# Sementi): se una di queste non compare qui, Groq non la usera' mai e
# la relativa spesa finira' accorpata genericamente in "Servizi" o
# "Materie Prime e Merci" invece di andare nella riga specifica.
USCITE_CATEGORIE: List[str] = [
    "Materie Prime e Merci",
    "Mangimi e Foraggi",
    "Carburanti",
    "Sementi",
    "Altri",
    "Servizi",
    "Energia elettrica",
    "Manutenzioni",
    "Lavorazioni c/terzi",
    "Leasing",
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
# 2. TEMPLATE EXCEL PREDEFINITO (fisso, incorporato nell'app)
# ---------------------------------------------------------------------------
# Il file "template_bilancio.xlsx" e' distribuito insieme al codice
# dell'app (stessa cartella): l'utente non deve piu' caricarlo ogni volta.
# La struttura del foglio (righe delle categorie, colonne degli anni) viene
# rilevata dinamicamente da excel_writer.py leggendo etichette e valori
# calcolati, quindi funziona anche se in futuro il template viene
# aggiornato con righe/colonne diverse.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template_bilancio.xlsx")

# Nome del foglio su cui si trovano le voci di riclassificazione (usato
# come preferenza; se non esiste, l'app usa comunque il foglio attivo).
SHEET_NAME_RICLASSIFICAZIONE = "Riclassificazione"

# Elenco anni di fallback, usato SOLO se il rilevamento automatico degli
# anni dal template fallisce per qualche motivo.
ANNI_DISPONIBILI: List[str] = ["2024", "2025"]


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

    return os.environ.get(key, default)


# Istanza di configurazione usata dall'applicazione.
GROQ_CONFIG = GroqConfig()


# ---------------------------------------------------------------------------
# 4. PALETTE COLORI (interfaccia utente)
# ---------------------------------------------------------------------------
PALETTE: Dict[str, str] = {
    "antracite": "#333333",     # Intestazioni principali, blocco finale
    "salvia": "#99B0A3",        # Elementi di spicco, accenti, macro-sezioni
    "tortora": "#D7D4D1",       # Bordi, tono medio
    "grigio_chiaro": "#E7E9E2", # Sfondi neutri delle card
    "crema": "#FDFBF7",         # Sfondo pagina
    "grigio_sfumato": "#787878",# Testo secondario
    "ocra": "#B87333",          # Accento caldo (hover, evidenze)
}
