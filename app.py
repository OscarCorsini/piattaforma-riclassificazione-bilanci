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

FLUSSO DI REVISIONE MANUALE (voce per voce):
---------------------------------------------
La classificazione automatica (parole chiave) resta come SUGGERIMENTO, ma
non viene piu' applicata direttamente: dopo l'estrazione, ogni singola voce
del bilancio viene mostrata all'utente in un popup con la categoria
suggerita gia' pre-selezionata. L'utente conferma o corregge, ed
eventualmente puo' assegnare un'etichetta personalizzata (max 20 caratteri)
al posto del nome della categoria nel foglio di dettaglio. Solo dopo aver
confermato tutte le voci viene generato il file Excel finale. Questo
elimina qualunque errore di classificazione automatica silenzioso: ogni
importo che finisce in una cella e' stato validato da una persona.

Avvio:
    streamlit run app.py

Prerequisiti / credenziali:
    Vedi commenti in config.py e groq_client.py per come impostare
    GROQ_API_KEY.
"""

from __future__ import annotations

import streamlit as st

from config import ANNI_DISPONIBILI, PALETTE, ENTRATE_CATEGORIE, USCITE_CATEGORIE
from pdf_extractor import extract_text_from_multiple_pdfs, PDFExtractionError
from groq_client import (
    estrai_voci_bilancio,
    GroqConfigError,
    GroqResponseError,
    Voce,
    RiclassificazioneResult,
)
from classificatore import suggerisci_categoria
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class^="st-emotion"], [class*=" st-emotion"],
    button, input, select, textarea, p, li, label, div,
    h1, h2, h3, h4, h5, h6 {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif !important;
    }}
    /* Le icone (es. quella di upload) sono glifi di un font-icona
       (Material Symbols): non vanno toccate, altrimenti si vede il
       testo "upload" al posto dell'icona. Escludere lo span dalla regola
       sopra NON basta, perche' font-family e' una proprieta' ereditata:
       lo span erediterebbe comunque "Inter" dal contenitore <div>
       genitore. Bisogna quindi RIPRISTINARE esplicitamente il font-icona
       sull'elemento stesso, per interrompere l'ereditarieta'. */
    span {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif !important;
    }}
    [data-testid="stIconMaterial"] {{
        font-family: 'Material Symbols Rounded' !important;
    }}

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
       allineamento verticale (le etichette sono sulla stessa riga, i
       campi sotto devono avere ESATTAMENTE la stessa altezza). */
    div[class*="st-key-riga_"] [data-testid="stHorizontalBlock"] {{
        align-items: flex-end;
    }}

    div[class*="st-key-riga_"] [data-testid="stWidgetLabel"] {{
        height: 24px;
        display: flex;
        align-items: center;
    }}

    [data-testid="stFileUploaderDropzone"] {{
        background-color: #ffffff;
        border: 1.5px dashed var(--tortora);
        border-radius: 8px;
        min-height: 44px;
        padding: 0.35rem 0.7rem;
        margin: 0;
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        box-sizing: border-box;
    }}
    /* Nasconde il testo informativo ("Drag and drop", "200MB per file")
       per rendere il campo compatto quanto il menu a tendina Anno quando
       e' vuoto. NON forziamo un'altezza massima fissa sul contenitore:
       quando un file e' stato caricato, Streamlit mostra un "chip" con
       nome/dimensione del file all'interno dello stesso contenitore, e
       un'altezza massima troppo rigida lo taglierebbe/sovrapporrebbe
       (il glitch di testo visto in precedenza). Il contenitore cresce
       quindi in altezza solo quando serve.
    */
    [data-testid="stFileUploaderDropzoneInstructions"] {{
        display: none;
    }}
    /* Il pulsante "Browse files" nativo non serve piu' dato che le
       istruzioni sono nascoste e vogliamo un campo compatto: lo teniamo
       ma senza farlo traboccare fuori dal contenitore. */
    [data-testid="stFileUploaderDropzone"] section > button {{
        margin-left: auto;
    }}
    /* Elenco dei file gia' caricati (chip con nome/dimensione): deve
       poter andare a capo ed avere spazio proprio, senza sovrapporsi
       al resto del contenuto del dropzone. */
    [data-testid="stFileUploaderDropzone"] ul,
    [data-testid="stFileUploaderFile"] {{
        width: 100%;
    }}

    div[data-baseweb="select"] {{
        height: 44px;
    }}
    div[data-baseweb="select"] > div {{
        border-radius: 6px;
        border-color: var(--tortora);
        height: 44px;
        min-height: 44px;
        max-height: 44px;
        box-sizing: border-box;
        display: flex;
        align-items: center;
    }}

    /* Il bottone "rimuovi riga" non ha un'etichetta sopra come gli altri
       due campi: lo spostiamo in basso per allinearlo alla stessa base. */
    div[class*="st-key-riga_"] div.stButton {{
        margin-top: 1.6rem;
    }}
    div[class*="st-key-riga_"] div.stButton > button {{
        height: 44px;
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
        <h1>Piattaforma di Riclassificazione Bilancio</h1>
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
# STATO WIZARD DI REVISIONE VOCE-PER-VOCE
# ---------------------------------------------------------------------------
def _reset_wizard():
    st.session_state.wiz_fase = "idle"
    st.session_state.wiz_voci = []
    st.session_state.wiz_indice = 0
    st.session_state.wiz_confermate = []
    st.session_state.wiz_gestione_fiscale = {}
    st.session_state.wiz_note = {}
    st.session_state.wiz_errori = {}
    st.session_state.wiz_buffer = None
    st.session_state.wiz_report = None
    st.session_state.wiz_errori_scrittura = None


if "wiz_fase" not in st.session_state:
    _reset_wizard()


def _formatta_euro(valore: float) -> str:
    testo = f"{valore:,.2f}"
    testo = testo.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"€ {testo}"


@st.dialog("Conferma voce estratta")
def _popup_revisione_voce():
    idx = st.session_state.wiz_indice
    voci = st.session_state.wiz_voci
    voce = voci[idx]

    st.caption(f"Voce {idx + 1} di {len(voci)} — Anno {voce['anno']}")
    sezione = "Ricavo" if voce["tipo"] == "entrata" else "Costo"
    st.markdown(f"**{sezione} rilevato dal bilancio:** {voce['descrizione']}")
    st.markdown(f"**Importo:** {_formatta_euro(voce['importo'])}")

    categorie = ENTRATE_CATEGORIE if voce["tipo"] == "entrata" else USCITE_CATEGORIE
    default_index = (
        categorie.index(voce["categoria_suggerita"])
        if voce["categoria_suggerita"] in categorie
        else 0
    )
    scelta = st.selectbox("Allocazione (voce del template)", options=categorie, index=default_index)
    etichetta_custom = st.text_input(
        "Etichetta personalizzata (opzionale, max 20 caratteri)",
        max_chars=20,
        placeholder=scelta,
    )

    col_conferma, col_salta = st.columns(2)
    with col_conferma:
        if st.button("Conferma", type="primary", use_container_width=True):
            descrizione_finale = etichetta_custom.strip() or scelta
            st.session_state.wiz_confermate.append({
                "anno": voce["anno"],
                "categoria": scelta,
                "descrizione": descrizione_finale,
                "importo": voce["importo"],
                "tipo": voce["tipo"],
            })
            st.session_state.wiz_indice += 1
            st.rerun()
    with col_salta:
        if st.button("Salta questa voce", use_container_width=True):
            st.session_state.wiz_indice += 1
            st.rerun()


def _costruisci_risultati_per_anno() -> dict:
    """Trasforma le scelte confermate voce-per-voce in un
    RiclassificazioneResult per anno, pronto per essere scritto nel
    template Excel."""
    risultati_per_anno: dict = {}

    anni_coinvolti = set(c["anno"] for c in st.session_state.wiz_confermate)
    anni_coinvolti |= set(st.session_state.wiz_gestione_fiscale.keys())

    for anno in anni_coinvolti:
        entrate = {c: [] for c in ENTRATE_CATEGORIE}
        uscite = {c: [] for c in USCITE_CATEGORIE}

        for conferma in st.session_state.wiz_confermate:
            if conferma["anno"] != anno:
                continue
            voce_obj = Voce(descrizione=conferma["descrizione"], importo=conferma["importo"])
            if conferma["tipo"] == "entrata":
                entrate.setdefault(conferma["categoria"], []).append(voce_obj)
            else:
                uscite.setdefault(conferma["categoria"], []).append(voce_obj)

        gestione_fiscale_anno = st.session_state.wiz_gestione_fiscale.get(anno, {})
        gestione_fiscale = {
            categoria: [Voce(descrizione="(valore aggregato dal bilancio)", importo=valore)]
            for categoria, valore in gestione_fiscale_anno.items()
        }

        risultati_per_anno[anno] = RiclassificazioneResult(
            entrate=entrate,
            uscite=uscite,
            gestione_fiscale=gestione_fiscale,
            note=st.session_state.wiz_note.get(anno, []),
        )

    return risultati_per_anno


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
                st.button("X", key=f"del_{riga_id}", on_click=_rimuovi_riga, args=(riga_id,))
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
    st.subheader("Estrai le voci e confermale una per una")
    st.caption(
        "Per ogni voce di ricavo o costo rilevata nel bilancio verra' "
        "chiesta conferma della categoria (con un suggerimento gia' "
        "pre-selezionato): puoi accettarlo o correggerlo prima di generare "
        "l'Excel finale."
    )

    righe_valide = [(anno, files) for anno, files in righe_dati if files]
    pronto = len(righe_valide) > 0 and st.session_state.wiz_fase == "idle"

    avvia = st.button(
        "Avvia estrazione",
        disabled=not (len(righe_valide) > 0 and st.session_state.wiz_fase == "idle"),
        use_container_width=True,
    )

    if len(righe_valide) == 0:
        st.caption("Carica almeno un PDF per un anno per procedere.")


# ---------------------------------------------------------------------------
# FASE 1: ESTRAZIONE (chiama l'AI, prepara la coda di voci da revisionare)
# ---------------------------------------------------------------------------
if avvia:
    _reset_wizard()
    voci_totali = []
    gestione_fiscale_per_anno = {}
    note_per_anno = {}
    errori_gruppo = {}

    with st.status("Estrazione in corso...", expanded=True) as status:
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

            status.write(f"Anno {anno}: lettura voci dal bilancio...")
            try:
                dati = estrai_voci_bilancio(anno, testi)
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

            for voce_raw in dati["voci"]:
                descrizione = voce_raw.get("descrizione", "")
                tipo = str(voce_raw.get("tipo", "")).strip().lower()
                tipo_normalizzato = "entrata" if tipo.startswith("entrat") else "uscita"
                try:
                    importo = float(voce_raw.get("importo", 0.0))
                except (TypeError, ValueError):
                    importo = 0.0
                if not descrizione or importo == 0.0:
                    continue
                voci_totali.append({
                    "anno": anno,
                    "descrizione": descrizione,
                    "importo": importo,
                    "tipo": tipo_normalizzato,
                    "categoria_suggerita": suggerisci_categoria(descrizione, tipo_normalizzato),
                })

            gestione_fiscale_per_anno[anno] = dati["gestione_fiscale"]
            note_per_anno[anno] = dati["note"]

        if not voci_totali and not gestione_fiscale_per_anno:
            status.update(label="Nessun anno elaborato con successo", state="error")
            for anno, msg in errori_gruppo.items():
                st.error(f"**Anno {anno}:** {msg}")
            st.stop()

        status.update(
            label=f"Estrazione completata - {len(voci_totali)} voci da confermare",
            state="complete",
        )

    st.session_state.wiz_voci = voci_totali
    st.session_state.wiz_gestione_fiscale = gestione_fiscale_per_anno
    st.session_state.wiz_note = note_per_anno
    st.session_state.wiz_errori = errori_gruppo
    st.session_state.wiz_indice = 0
    st.session_state.wiz_confermate = []
    st.session_state.wiz_fase = "revisione" if voci_totali else "completato"
    st.rerun()


# ---------------------------------------------------------------------------
# FASE 2: REVISIONE (un popup per voce, in sequenza)
# ---------------------------------------------------------------------------
if st.session_state.wiz_fase == "revisione":
    idx = st.session_state.wiz_indice
    totale = len(st.session_state.wiz_voci)

    if idx < totale:
        st.progress(idx / totale, text=f"Revisione voce {idx + 1} di {totale}")
        _popup_revisione_voce()
    else:
        st.session_state.wiz_fase = "completato"
        st.rerun()


# ---------------------------------------------------------------------------
# FASE 3: COMPLETATO - genera l'Excel dalle scelte confermate
# ---------------------------------------------------------------------------
if st.session_state.wiz_fase == "completato":
    for anno, msg in st.session_state.wiz_errori.items():
        st.warning(f"Anno {anno} non elaborato: {msg}")

    if st.session_state.wiz_buffer is None:
        risultati_per_anno = _costruisci_risultati_per_anno()

        if not risultati_per_anno:
            st.error("Nessuna voce confermata: impossibile generare l'Excel.")
            st.button("Ricomincia", on_click=_reset_wizard, use_container_width=True)
            st.stop()

        try:
            buffer, report_per_anno, errori_scrittura = popola_template_multi(
                template_bytes, risultati_per_anno
            )
        except ExcelTemplateError as e:
            st.error(f"**Errore nel template Excel.**\n\n{e}")
            st.button("Ricomincia", on_click=_reset_wizard, use_container_width=True)
            st.stop()

        st.session_state.wiz_buffer = buffer
        st.session_state.wiz_report = report_per_anno
        st.session_state.wiz_errori_scrittura = errori_scrittura

    report_per_anno = st.session_state.wiz_report
    errori_scrittura = st.session_state.wiz_errori_scrittura

    for anno, msg in errori_scrittura.items():
        st.warning(f"Anno {anno} non scritto nel file: {msg}")

    for anno, report in report_per_anno.items():
        if report.voci_non_mappate:
            st.warning(
                f"Anno {anno}: le seguenti categorie non sono state trovate "
                "nel template e non sono state scritte: "
                + ", ".join(report.voci_non_mappate)
            )

    with st.expander("Dettaglio voci confermate"):
        dettagli = [
            {
                "Anno": c["anno"],
                "Sezione": "Entrate" if c["tipo"] == "entrata" else "Uscite",
                "Categoria": c["categoria"],
                "Etichetta": c["descrizione"],
                "Importo (EUR)": c["importo"],
            }
            for c in st.session_state.wiz_confermate
        ]
        if dettagli:
            st.dataframe(dettagli, use_container_width=True)
        else:
            st.info("Nessuna voce confermata.")

    totale_celle = sum(r.celle_scritte for r in report_per_anno.values())
    anni_ok = ", ".join(sorted(report_per_anno.keys()))
    st.success(
        f"Elaborazione completata per gli anni: {anni_ok} "
        f"({totale_celle} celle aggiornate)."
    )

    st.download_button(
        label="Scarica il file Excel completato",
        data=st.session_state.wiz_buffer,
        file_name="Riclassificazione_bilancio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.button("Nuova elaborazione", on_click=_reset_wizard, use_container_width=True)
