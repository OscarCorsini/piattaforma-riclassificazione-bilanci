# -*- coding: utf-8 -*-
"""
pdf_extractor.py
================
Estrazione del testo grezzo da uno o più PDF di situazione contabile
(bilancio di verifica, mastrini, estratti conto, ecc.).

Usa pdfplumber (più accurato con tabelle) con fallback su PyPDF2 se
pdfplumber non riesce a leggere il file (es. PDF scansionati senza layer
testuale, PDF corrotti, protetti da password non gestita, ecc.).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

import pdfplumber
from PyPDF2 import PdfReader


class PDFExtractionError(Exception):
    """Sollevata quando un PDF non può essere letto da nessuno dei due motori."""


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

                # Estrae anche le tabelle in forma testuale semplice, utile
                # perché le situazioni contabili sono spesso tabellari.
                for table in page.extract_tables():
                    for row in table:
                        clean_row = [str(c) if c is not None else "" for c in row]
                        parts.append(" | ".join(clean_row))
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
    Estrae il testo da più file PDF caricati contemporaneamente.
    Ogni file viene elaborato in modo indipendente: se uno fallisce,
    l'eccezione riporta chiaramente quale file ha causato il problema
    (il chiamante decide se interrompere o continuare con gli altri).
    """
    documents: List[ExtractedDocument] = []
    for f in files:
        doc = extract_text_from_pdf(f, filename=getattr(f, "name", "documento.pdf"))
        documents.append(doc)
    return documents
