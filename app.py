# -*- coding: utf-8 -*-
"""
app.py
======
Interfaccia Streamlit per l'automazione della riclassificazione finanziaria
di bilanci/situazioni contabili aziendali (settore agricolo), con output
finale in Excel.

Il template Excel e' fisso e incorporato nell'app (template_bilancio.xlsx):
l'utente non deve piu' caricarlo manualmente. E' possibile elaborare piu'
anni nella stessa sessione: per ciascun anno si sceglie il periodo da un
menu a tendina e si caricano i PDF corrispondenti; tutti gli anni vengono
scritti in un unico file Excel scaricabile.

Avvio:
    streamlit run app.py

Prerequisiti / credenziali:
    Vedi commenti in config.py e groq_client.py per come impostare
    GROQ_API_KEY.
"""

from __future__ import annotations

import streamlit as st

from config import ANNI_DISPONIBILI, PALETTE
from pdf_extractor import extract_text_from_multiple_pdfs, PDFExtractionError
from groq_client import (
    riclassifica_bilancio,
    GroqConfigError,
    GroqResponseError,
)
from excel_writer import (
    carica_template_predefinito,
    rileva_anni_disponibili,
    popola_template_multi,
    ExcelTemplateError,
)


# ---------------------------------------------------------------------------
# CONFIGURAZIONE PAGINA E STILE
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Riclassificazione Finanziaria | Piattaforma Agritech",
    page_icon="\U0001F4CA",
    layout="centered",
    initial_sidebar_state="collapsed",
)

CUSTOM_CSS = f"""
<style>
    :root {{
        --antracite: {PALETTE['antracite']};
        --salvia: {PALETTE['salvia']};
        --tortora: {PALETTE['tortora']};
        --grigio-chiaro: {PALETTE['grigio_chiaro']};
        --crema: {PALETTE['crema']};
        --grigio-sfumato: {PALETTE['grigio_sfumato']};
        --ocra: {PALETTE['ocra']};
    }}

    #MainMenu, footer {{visibility: hidden;}}

    [data-testid="stAppViewContainer"] {{
        background-color: var(--crema);
    }}

    .block-container {{
        max-width: 820px;
        padding-top: 2.6rem;
        padding-bottom: 3rem;
    }}

    h1, h2, h3 {{
        color: var(--antracite);
        font-weight: 600;
        letter-spacing: -0.01em;
    }}

    p, li, label {{
        color: var(--antracite);
    }}

    /* Il testo dei bottoni Streamlit e' avvolto in <p>/<span> interni: la
       regola generica sopra lo renderebbe scuro su sfondo scuro. Qui lo
       forziamo sempre bianco per i bottoni con sfondo pieno. */
    div.stButton > button p,
    div.stButton > button span,
    div.stButton > button div,
    div[data-testid="stDownloadButton"] > button p,
    div[data-testid="stDownloadButton"] > button span,
    div[data-testid="stDownloadButton"] > button div {{
        color: inherit !important;
    }}

    .app-header {{
        margin-bottom: 1.8rem;
    }}

    .app-subtitle {{
        color: var(--grigio-sfumato);
        font-size: 0.96rem;
        margin-top: 0.2rem;
    }}

    /* Card reale (Streamlit st.container(border=True)), non un div HTML
       "manuale" che rimarrebbe vuoto tra una chiamata st.markdown e
       l'altra. */
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background: #ffffff;
        border: 1px solid var(--tortora) !important;
        border-left: 4px solid var(--salvia) !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 3px rgba(51, 51, 51, 0.05);
    }}

    .step-label {{
        display: inline-block;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #ffffff;
        background-color: var(--antracite);
        font-weight: 700;
        padding: 0.18rem 0.6rem;
        border-radius: 5px;
        margin-bottom: 0.6rem;
    }}

    div[class*="st-key-riga_"] {{
        background: var(--grigio-chiaro);
        border: 1px solid var(--tortora);
        border-radius: 10px;
        padding: 0.9rem 1rem 0.9rem 1rem;
        margin-bottom: 0.8rem;
    }}

    .anni-info {{
        color: var(--grigio-sfumato);
        font-size: 0.88rem;
        font-style: italic;
    }}

    div.stButton > button {{
        background-color: var(--antracite);
        color: #ffffff;
        border: none;
        border-radius: 6px;
        padding: 0.55rem 1.4rem;
        font-weight: 500;
        width: 100%;
        transition: background-color 0.15s ease;
    }}
    div.stButton > button:hover {{
        background-color: var(--ocra);
        color: #ffffff;
    }}

    div[data-testid="stDownloadButton"] > button {{
        background-color: var(--salvia);
        color: var(--antracite);
        border: none;
        border-radius: 6px;
        width: 100%;
        font-weight: 600;
    }}
    div[data-testid="stDownloadButton"] > button:hover {{
        background-color: var(--ocra);
        color: #ffffff;
    }}

    /* Allineamento riga Anno / PDF / rimuovi: stessa altezza e stesso
       allineamento verticale. */
    div[class*="st-key-riga_"] [data-testid="stHorizontalBlock"] {{
        align-items: flex-end;
    }}

    [data-testid="stFileUploaderDropzone"] {{
        background-color: #ffffff;
        border: 1.5px dashed var(--tortora);
        border-radius: 8px;
        min-height: 40px;
        padding: 0.35rem 0.7rem;
        display: flex;
        align-items: center;
    }}
    /* Nasconde il testo informativo ("Drag and drop", "200MB per file")
       per rendere il campo compatto quanto il menu a tendina Anno. */
    [data-testid="stFileUploaderDropzoneInstructions"] {{
        display: none;
    }}

    div[data-baseweb="select"] > div {{
        border-radius: 6px;
        border-color: var(--tortora);
        min-height: 40px;
    }}

    /* Il bottone "rimuovi riga" non ha un'etichetta sopra come gli altri
       due campi: lo spostiamo in basso per allinearlo alla stessa base. */
    div[class*="st-key-riga_"] div.stButton {{
        margin-top: 1.6rem;
    }}

    footer {{display: none;}}
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
            Analisi delle situazioni contabili e generazione del modello
            di stima del margine.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# CARICAMENTO TEMPLATE PREDEFINITO (fisso, incorporato nell'app)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _carica_template_e_anni():
    dati = carica_template_predefinito()
    anni = rileva_anni_disponibili(dati)
    return dati, anni


try:
    template_bytes, anni_rilevati = _carica_template_e_anni()
except ExcelTemplateError as e:
    st.error(f"**Impossibile caricare il template predefinito.**\n\n{e}")
    st.stop()

opzioni_anni = anni_rilevati if anni_rilevati else ANNI_DISPONIBILI
if not anni_rilevati:
    st.warning(
        "Non e' stato possibile rilevare automaticamente gli anni dal "
        "template predefinito: uso l'elenco di default."
    )


# ---------------------------------------------------------------------------
# STATO: RIGHE ANNO + PDF
# ---------------------------------------------------------------------------
if "righe_ids" not in st.session_state:
    st.session_state.righe_ids = [0]
if "prossimo_id" not in st.session_state:
    st.session_state.prossimo_id = 1


def _aggiungi_riga():
    st.session_state.righe_ids.append(st.session_state.prossimo_id)
    st.session_state.prossimo_id += 1


def _rimuovi_riga(riga_id: int):
    if riga_id in st.session_state.righe_ids:
        st.session_state.righe_ids.remove(riga_id)
    if not st.session_state.righe_ids:
        st.session_state.righe_ids = [st.session_state.prossimo_id]
        st.session_state.prossimo_id += 1


# ---------------------------------------------------------------------------
# STEP UNICO - SELEZIONE ANNI E CARICAMENTO PDF
# ---------------------------------------------------------------------------
with st.container(border=True):
    st.markdown('<div class="step-label">Situazioni contabili</div>', unsafe_allow_html=True)
    st.subheader("Seleziona l'anno e carica i documenti")
    st.caption(
        "Il modello Excel e' quello aziendale standard: non serve caricarlo. "
        "Puoi elaborare piu' anni nella stessa esecuzione: aggiungi una riga "
        "per ciascun anno da compilare."
    )

    righe_dati = []
    for riga_id in list(st.session_state.righe_ids):
        with st.container(key=f"riga_{riga_id}"):
            col_anno, col_upload, col_rimuovi = st.columns([2, 5, 1])
            with col_anno:
                anno_scelto = st.selectbox(
                    "Anno",
                    options=opzioni_anni,
                    key=f"anno_{riga_id}",
                )
            with col_upload:
                file_pdf = st.file_uploader(
                    "PDF situazione contabile",
                    type=["pdf"],
                    accept_multiple_files=True,
                    key=f"pdf_{riga_id}",
                )
            with col_rimuovi:
                st.button("✕", key=f"del_{riga_id}", on_click=_rimuovi_riga, args=(riga_id,))
        righe_dati.append((anno_scelto, file_pdf))

    col_add, col_info = st.columns([2, 5])
    with col_add:
        st.button("+ Aggiungi un altro anno", on_click=_aggiungi_riga, use_container_width=True)
    with col_info:
        st.markdown(
            f'<div class="anni-info">Anni disponibili nel template: {", ".join(opzioni_anni)}</div>',
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# STEP FINALE - ELABORAZIONE
# ---------------------------------------------------------------------------
with st.container(border=True):
    st.markdown('<div class="step-label">Elaborazione</div>', unsafe_allow_html=True)
    st.subheader("Genera il file Excel completato")

    righe_valide = [(anno, files) for anno, files in righe_dati if files]
    pronto = len(righe_valide) > 0

    avvia = st.button("Avvia elaborazione", disabled=not pronto, use_container_width=True)

    if not pronto:
        st.caption("Carica almeno un PDF per un anno per procedere.")


# ---------------------------------------------------------------------------
# LOGICA DI ELABORAZIONE
# ---------------------------------------------------------------------------
if avvia:
    risultati_per_anno = {}
    errori_gruppo = {}

    with st.status("Elaborazione in corso...", expanded=True) as status:
        for anno, files in righe_valide:
            status.write(f"Anno {anno}: estrazione testo dai PDF...")
            try:
                documenti = extract_text_from_multiple_pdfs(files)
            except PDFExtractionError as e:
                errori_gruppo[anno] = f"Errore di lettura PDF: {e}"
                continue

            testi = [d.text for d in documenti]
            for doc in documenti:
                for w in doc.warnings:
                    st.warning(f"Anno {anno} - '{doc.filename}': {w}")

            status.write(f"Anno {anno}: riclassificazione in corso...")
            try:
                risultato = riclassifica_bilancio(anno, testi)
            except GroqConfigError as e:
                status.update(label="Configurazione mancante", state="error")
                st.error(
                    "**Il motore di analisi non e' configurato.**\n\n"
                    f"{e}"
                )
                st.stop()
            except GroqResponseError as e:
                errori_gruppo[anno] = f"Risposta non nel formato atteso: {e}"
                continue

            risultati_per_anno[anno] = risultato
            if risultato.note:
                with st.expander(f"Osservazioni - anno {anno}"):
                    for nota in risultato.note:
                        st.write(f"- {nota}")

        if not risultati_per_anno:
            status.update(label="Nessun anno elaborato con successo", state="error")
            for anno, msg in errori_gruppo.items():
                st.error(f"**Anno {anno}:** {msg}")
            st.stop()

        status.write("Popolamento del template Excel...")
        try:
            buffer, report_per_anno, errori_scrittura = popola_template_multi(
                template_bytes, risultati_per_anno
            )
        except ExcelTemplateError as e:
            status.update(label="Errore nel file Excel", state="error")
            st.error(f"**Errore nel template Excel.**\n\n{e}")
            st.stop()

        for anno, msg in errori_gruppo.items():
            st.warning(f"Anno {anno} non elaborato: {msg}")
        for anno, msg in errori_scrittura.items():
            st.warning(f"Anno {anno} non scritto nel file: {msg}")

        for anno, report in report_per_anno.items():
            if report.voci_non_mappate:
                st.warning(
                    f"Anno {anno}: le seguenti categorie non sono state trovate "
                    "nel template e non sono state scritte: "
                    + ", ".join(report.voci_non_mappate)
                )

        totale_celle = sum(r.celle_scritte for r in report_per_anno.values())
        status.update(
            label=f"Completato - {totale_celle} celle aggiornate su {len(report_per_anno)} anni",
            state="complete",
        )

    anni_ok = ", ".join(sorted(report_per_anno.keys()))
    st.success(f"Elaborazione completata per gli anni: {anni_ok}.")

    st.download_button(
        label="Scarica il file Excel completato",
        data=buffer,
        file_name="Riclassificazione_bilancio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
