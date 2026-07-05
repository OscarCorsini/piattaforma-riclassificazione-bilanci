# -*- coding: utf-8 -*-
"""
app.py
======
Interfaccia Streamlit per l'automazione della riclassificazione finanziaria
di bilanci/situazioni contabili aziendali (settore agricolo), tramite
Google Gemini come motore AI, con output finale in Excel.

Avvio:
    streamlit run app.py

Prerequisiti / credenziali:
    Vedi commenti in config.py e gemini_client.py per come impostare
    GEMINI_API_KEY.
"""

from __future__ import annotations

import datetime as _dt

import streamlit as st

from config import ANNI_DISPONIBILI
from pdf_extractor import extract_text_from_multiple_pdfs, PDFExtractionError
from gemini_client import (
    riclassifica_bilancio,
    GeminiConfigError,
    GeminiResponseError,
)
from excel_writer import popola_template_excel, ExcelTemplateError, ExcelMappingError


# ---------------------------------------------------------------------------
# CONFIGURAZIONE PAGINA E STILE
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Riclassificazione Finanziaria | Piattaforma Agritech",
    page_icon="\U0001F4CA",
    layout="centered",
    initial_sidebar_state="collapsed",
)

CUSTOM_CSS = """
<style>
    /* Palette neutra e professionale */
    :root {
        --brand-dark: #1a2b3c;
        --brand-accent: #2f6f4e;
        --surface: #ffffff;
        --border-color: #e3e7eb;
    }

    #MainMenu, footer {visibility: hidden;}

    .block-container {
        max-width: 780px;
        padding-top: 2.5rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3 {
        color: var(--brand-dark);
        font-weight: 600;
        letter-spacing: -0.01em;
    }

    .app-header {
        border-bottom: 1px solid var(--border-color);
        padding-bottom: 1.2rem;
        margin-bottom: 2rem;
    }

    .app-subtitle {
        color: #5b6b7a;
        font-size: 0.95rem;
        margin-top: -0.6rem;
    }

    .section-card {
        background: var(--surface);
        border: 1px solid var(--border-color);
        border-radius: 10px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1.4rem;
    }

    .step-label {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: var(--brand-accent);
        font-weight: 700;
        margin-bottom: 0.3rem;
    }

    div.stButton > button {
        background-color: var(--brand-dark);
        color: #ffffff;
        border: none;
        border-radius: 6px;
        padding: 0.55rem 1.4rem;
        font-weight: 500;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: var(--brand-accent);
        color: #ffffff;
    }

    div[data-testid="stDownloadButton"] > button {
        background-color: var(--brand-accent);
        color: #ffffff;
        border: none;
        border-radius: 6px;
        width: 100%;
        font-weight: 500;
    }

    footer {display: none;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# INTESTAZIONE
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="app-header">
        <h1>Piattaforma di Riclassificazione Finanziaria</h1>
        <p class="app-subtitle">
            Analisi automatica delle situazioni contabili con Google Gemini
            e generazione del modello di stima del margine in Excel.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# STEP 1 - CARICAMENTO TEMPLATE
# ---------------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="step-label">Passo 1</div>', unsafe_allow_html=True)
st.subheader("Carica il modello Excel (Template)")

template_file = st.file_uploader(
    "File Excel del modello di stima del margine",
    type=["xlsx"],
    key="template_uploader",
    help="Il file non verra' modificato: verra' generata una copia con i dati popolati.",
)
st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# STEP 2 - ANNO DI RIFERIMENTO
# ---------------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="step-label">Passo 2</div>', unsafe_allow_html=True)
st.subheader("Seleziona l'anno di riferimento")

anno_corrente = str(_dt.date.today().year)
default_index = ANNI_DISPONIBILI.index(anno_corrente) if anno_corrente in ANNI_DISPONIBILI else 0
anno_selezionato = st.selectbox(
    "Anno della situazione contabile",
    options=ANNI_DISPONIBILI,
    index=default_index,
    key="anno_select",
)
st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# STEP 3 - CARICAMENTO PDF
# ---------------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="step-label">Passo 3</div>', unsafe_allow_html=True)
st.subheader("Carica la situazione contabile (PDF)")

pdf_files = st.file_uploader(
    "Uno o piu' file PDF (bilancio di verifica, mastrini, situazione contabile)",
    type=["pdf"],
    accept_multiple_files=True,
    key="pdf_uploader",
)
st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# STEP 4 - ELABORAZIONE
# ---------------------------------------------------------------------------
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="step-label">Passo 4</div>', unsafe_allow_html=True)
st.subheader("Elabora e genera il file finale")

pronto = template_file is not None and pdf_files and len(pdf_files) > 0

avvia = st.button("Avvia elaborazione", disabled=not pronto, use_container_width=True)

if not pronto:
    st.caption("Carica il template Excel e almeno un PDF per procedere.")

st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# LOGICA DI ELABORAZIONE
# ---------------------------------------------------------------------------
if avvia:
    try:
        with st.status("Elaborazione in corso...", expanded=True) as status:

            # --- Fase 1: estrazione testo dai PDF ---
            status.write("Estrazione testo dai documenti PDF...")
            try:
                documenti = extract_text_from_multiple_pdfs(pdf_files)
            except PDFExtractionError as e:
                status.update(label="Errore nella lettura del PDF", state="error")
                st.error(f"**Errore di lettura PDF**\n\n{e}")
                st.stop()

            testi = [d.text for d in documenti]
            for doc in documenti:
                for w in doc.warnings:
                    st.warning(f"'{doc.filename}': {w}")

            # --- Fase 2: chiamata a Gemini per la riclassificazione ---
            status.write("Interpretazione e riclassificazione con Google Gemini...")
            try:
                risultato = riclassifica_bilancio(anno_selezionato, testi)
            except GeminiConfigError as e:
                status.update(label="Configurazione Gemini mancante", state="error")
                st.error(
                    "**Gemini non e' configurato.**\n\n"
                    f"{e}\n\n"
                    "Genera una chiave gratuita su https://aistudio.google.com/apikey "
                    "e impostala come GEMINI_API_KEY nei Secrets dell'app."
                )
                st.stop()
            except GeminiResponseError as e:
                status.update(label="Risposta Gemini non valida", state="error")
                st.error(f"**Risposta AI non nel formato atteso.**\n\n{e}")
                st.stop()

            if risultato.note:
                with st.expander("Osservazioni dell'AI sulle voci elaborate"):
                    for nota in risultato.note:
                        st.write(f"- {nota}")

            # --- Fase 3: scrittura nel template Excel ---
            status.write("Popolamento del template Excel...")
            try:
                buffer, report = popola_template_excel(template_file, anno_selezionato, risultato)
            except ExcelTemplateError as e:
                status.update(label="Errore nel file Excel", state="error")
                st.error(f"**Errore nel template Excel.**\n\n{e}")
                st.stop()
            except ExcelMappingError as e:
                status.update(label="Mapping anno non configurato", state="error")
                st.error(f"**Configurazione mancante.**\n\n{e}")
                st.stop()

            if report.voci_non_mappate:
                st.warning(
                    "Le seguenti categorie non hanno una cella configurata per "
                    f"l'anno {anno_selezionato} e non sono state scritte nel file: "
                    + ", ".join(report.voci_non_mappate)
                )

            status.update(
                label=f"Completato - {report.celle_scritte} celle aggiornate",
                state="complete",
            )

        st.success(
            f"Elaborazione completata. {report.celle_scritte} valori scritti nel "
            f"foglio '{report.foglio}' per l'anno {anno_selezionato}."
        )

        nome_file_output = f"Riclassificazione_{anno_selezionato}.xlsx"
        st.download_button(
            label="Scarica il file Excel completato",
            data=buffer,
            file_name=nome_file_output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        # Rete di sicurezza per qualsiasi errore imprevisto non gia' gestito
        # dai blocchi specifici sopra: mostra un messaggio chiaro invece di
        # un traceback grezzo all'utente finale.
        st.error(
            "**Si e' verificato un errore imprevisto durante l'elaborazione.**\n\n"
            f"Dettaglio tecnico: `{type(e).__name__}: {e}`"
        )
