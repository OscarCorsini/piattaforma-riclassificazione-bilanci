# -*- coding: utf-8 -*-
"""
pdf_extractor.py
================
Estrazione del testo grezzo da uno o piu' PDF di situazione contabile
(bilancio di verifica, mastrini, estratti conto, ecc.).

Usa pdfplumber (piu' accurato con tabelle) con fallback su PyPDF2 se
pdfplumber non riesce a leggere il file (es. PDF scansionati senza layer
testuale, PDF corrotti, protetti da password non gestita, ecc.).
"""

from dataclasses import dataclass
from typing import List, Optional

import pdfplumber
from PyPDF2 import PdfReader


class PDFExtractionError(Exception):
    """Sollevata quando un PDF non puo' essere letto da nessuno dei due motori."""


@dataclass
class ExtractedDocument:
    filename: str
    text: str
    n_pages: int
    warnings: List[str]


def _extract_with_pdfplumber(file_obj) -> Optional[str]:
    try:
        file_obj.seek(0)
        parts = []
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                parts.append(page_text)

                # Estrae anche le tabelle in forma testuale semplice, ma
                # SOLO come ripiego per le pagine dove extract_text() non ha
                # restituito contenuto utile (es. PDF con layout tabellare
                # complesso che confonde l'estrazione lineare). Se il testo
                # lineare e' gia' presente, NON duplicare le stesse voci
                # anche come dump di tabella: per bilanci con molte righe
                # questo raddoppiava (e sporcava con celle vuote separate da
                # "|") il testo inviato all'AI, rendendo l'estrazione delle
                # singole voci molto meno affidabile.
                if len(page_text.strip()) < 40:
                    for table in page.extract_tables():
                        for row in table:
                            clean_row = [str(c) if c is not None else "" for c in row]
                            riga = " | ".join(clean_row).strip(" |")
                            if riga:
                                parts.append(riga)
        text = "\n".join(parts).strip()
        return text if text else None
    except Exception:
        return None


def _extract_with_pypdf2(file_obj) -> Optional[str]:
    try:
        file_obj.seek(0)
        reader = PdfReader(file_obj)
        parts = [page.extract_text() or "" for page in reader.pages]
        text = "\n".join(parts).strip()
        return text if text else None
    except Exception:
        return None


def extract_text_from_pdf(file_obj, filename: str) -> ExtractedDocument:
    """
    Estrae il testo da un file PDF (file-like object, es. da
    st.file_uploader). Tenta pdfplumber, poi PyPDF2 come fallback.

    Solleva PDFExtractionError con un messaggio chiaro se nessuno dei due
    motori riesce a estrarre contenuto testuale utile (es. scansione senza
    OCR).
    """
    warnings: List[str] = []

    text = _extract_with_pdfplumber(file_obj)
    if text is None:
        warnings.append(
            "Estrazione con pdfplumber non riuscita: utilizzato motore di fallback (PyPDF2)."
        )
        text = _extract_with_pypdf2(file_obj)

    if text is None or len(text.strip()) < 20:
        raise PDFExtractionError(
            f"Impossibile estrarre testo leggibile dal file '{filename}'. "
            "Il PDF potrebbe essere una scansione immagine priva di OCR, "
            "protetto da password, oppure danneggiato. "
            "Verifica il file oppure fornisci una versione con testo selezionabile."
        )

    try:
        file_obj.seek(0)
        n_pages = len(PdfReader(file_obj).pages)
    except Exception:
        n_pages = 0

    return ExtractedDocument(filename=filename, text=text, n_pages=n_pages, warnings=warnings)


def extract_text_from_multiple_pdfs(files) -> List[ExtractedDocument]:
    """
    Estrae il testo da piu' file PDF caricati contemporaneamente.
    Ogni file viene elaborato in modo indipendente: se uno fallisce,
    l'eccezione riporta chiaramente quale file ha causato il problema
    (il chiamante decide se interrompere o continuare con gli altri).
    """
    documents: List[ExtractedDocument] = []
    for f in files:
        doc = extract_text_from_pdf(f, filename=getattr(f, "name", "documento.pdf"))
        documents.append(doc)
    return documents
