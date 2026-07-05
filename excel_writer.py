# -*- coding: utf-8 -*-
"""
excel_writer.py
===============
Scrittura dei dati riclassificati nel file Excel "Template" tramite
openpyxl, senza alterare formule o formattazioni preesistenti.

Il modulo rileva DINAMICAMENTE la struttura del template (righe delle
categorie, colonne degli anni) invece di usare un mapping statico, cosi'
funziona con template reali che possono differire da quello di esempio:

  1. Trova la riga di ciascuna categoria cercando l'etichetta testuale
     nel foglio (es. "Corrispettivi normali").
  2. Rileva la colonna corrispondente a ciascun anno leggendo i valori
     calcolati da Excel nella riga di intestazione degli anni (utile
     quando l'anno e' il risultato di una formula, es. "anno base + 1").
  3. Se la cella di destinazione fa parte di un intervallo di celle unite
     (merged), scrive nella cella "ancora" (in alto a sinistra).
  4. Se una categoria non viene trovata nel template, la segnala come
     "non mappata" invece di generare un errore bloccante.

Supporta inoltre la scrittura di PIU' ANNI in un'unica esecuzione
(popola_template_multi), aprendo il workbook una sola volta e scrivendo
tutti gli anni richiesti prima di salvare.
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from config import SHEET_NAME_RICLASSIFICAZIONE, TEMPLATE_PATH
from groq_client import RiclassificazioneResult


class ExcelTemplateError(Exception):
    """Errore nell'apertura o nella struttura del file Excel template."""


class ExcelMappingError(Exception):
    """Errore nel mapping tra l'anno selezionato e le colonne del template."""


@dataclass
class WriteReport:
    """Riepilogo di cio' che e' stato scritto, utile per un feedback chiaro in UI."""
    celle_scritte: int
    voci_non_mappate: List[str]
    foglio: str


# ---------------------------------------------------------------------------
# UTILITA' DI LETTURA / RICERCA NEL FOGLIO
# ---------------------------------------------------------------------------
def _get_sheet(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    return workbook.active


def _normalizza_etichetta(valore) -> str:
    if valore is None:
        return ""
    testo = str(valore)
    testo = unicodedata.normalize("NFKD", testo)
    testo = testo.encode("ascii", "ignore").decode("ascii")
    testo = re.sub(r"[^a-zA-Z0-9]+", " ", testo)
    return re.sub(r"\s+", " ", testo).strip().lower()


def _trova_riga_categoria(sheet, categoria: str, max_row: int = 300, max_col: int = 6) -> Optional[int]:
    target = _normalizza_etichetta(categoria)
    if not target:
        return None
    limite_righe = min(sheet.max_row or max_row, max_row)
    limite_colonne = min(sheet.max_column or max_col, max_col)
    for row in range(1, limite_righe + 1):
        for col in range(1, limite_colonne + 1):
            valore = sheet.cell(row=row, column=col).value
            if valore is not None and _normalizza_etichetta(valore) == target:
                return row
    return None


def _rileva_colonne_anno(sheet_valori, max_row: int = 30, max_col: int = 40) -> Dict[str, str]:
    """Restituisce le colonne-anno {colonna_lettera: anno} della PRIMA riga
    (dall'alto) che contiene almeno 2 valori interi verosimili come anno."""
    limite_righe = min(sheet_valori.max_row or max_row, max_row)
    limite_colonne = min(sheet_valori.max_column or max_col, max_col)

    for row in range(1, limite_righe + 1):
        colonne_riga: Dict[str, str] = {}
        for col in range(1, limite_colonne + 1):
            valore = sheet_valori.cell(row=row, column=col).value
            if isinstance(valore, bool):
                continue
            if isinstance(valore, (int, float)):
                anno_int = int(valore)
                if 1900 <= anno_int <= 2100 and float(anno_int) == float(valore):
                    colonne_riga[get_column_letter(col)] = str(anno_int)
        if len(colonne_riga) >= 2:
            return colonne_riga

    return {}


def _trova_cella_ancora(sheet, cella_unita):
    for rng in sheet.merged_cells.ranges:
        if cella_unita.coordinate in rng:
            return sheet.cell(row=rng.min_row, column=rng.min_col)
    return None


def _leggi_bytes(file_like) -> bytes:
    try:
        file_like.seek(0)
    except Exception:
        pass
    dati = file_like.read()
    try:
        file_like.seek(0)
    except Exception:
        pass
    return dati


def _scrivi_valori_anno(sheet, colonna: str, risultato: RiclassificazioneResult) -> WriteReport:
    """Scrive i valori di UN anno (una colonna) gia' individuata nel foglio
    gia' aperto. Funzione interna riusata sia dal percorso a singolo anno
    che da quello multi-anno."""
    valori_da_scrivere = {
        **risultato.entrate,
        **risultato.uscite,
        **risultato.gestione_fiscale,
    }

    celle_scritte = 0
    voci_non_mappate: List[str] = []

    for categoria, valore in valori_da_scrivere.items():
        riga = _trova_riga_categoria(sheet, categoria)
        if riga is None:
            voci_non_mappate.append(categoria)
            continue

        cella = sheet[f"{colonna}{riga}"]
        if isinstance(cella, MergedCell):
            ancora = _trova_cella_ancora(sheet, cella)
            if ancora is None:
                voci_non_mappate.append(categoria)
                continue
            cella = ancora

        try:
            cella.value = valore
            celle_scritte += 1
        except Exception as exc:
            raise ExcelTemplateError(
                f"Errore scrivendo la categoria '{categoria}' nella cella '{cella.coordinate}': {exc}"
            ) from exc

    return WriteReport(
        celle_scritte=celle_scritte,
        voci_non_mappate=voci_non_mappate,
        foglio=sheet.title,
    )


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA: TEMPLATE PREDEFINITO
# ---------------------------------------------------------------------------
def carica_template_predefinito() -> bytes:
    """Legge dal disco il template Excel fisso distribuito con l'app."""
    try:
        with open(TEMPLATE_PATH, "rb") as f:
            return f.read()
    except Exception as exc:
        raise ExcelTemplateError(
            f"Impossibile leggere il template predefinito ({TEMPLATE_PATH}): {exc}"
        ) from exc


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA: RILEVAMENTO ANNI DISPONIBILI
# ---------------------------------------------------------------------------
def rileva_anni_disponibili(template_bytes: bytes) -> List[str]:
    try:
        workbook_valori = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=True)
        sheet_valori = _get_sheet(workbook_valori, SHEET_NAME_RICLASSIFICAZIONE)
        colonne_anno = _rileva_colonne_anno(sheet_valori)
        return sorted(set(colonne_anno.values()))
    except Exception:
        return []


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA: SCRITTURA MULTI-ANNO (usata dall'interfaccia)
# ---------------------------------------------------------------------------
def popola_template_multi(
    template_bytes: bytes,
    risultati_per_anno: Dict[str, RiclassificazioneResult],
) -> Tuple[io.BytesIO, Dict[str, WriteReport], Dict[str, str]]:
    """
    Apre il template UNA SOLA VOLTA e scrive i dati di piu' anni nello
    stesso file, restituendo un unico buffer scaricabile.

    Restituisce (buffer, report_per_anno, errori_per_anno): gli anni scritti
    con successo compaiono in report_per_anno, quelli per cui non e' stata
    trovata una colonna nel template compaiono in errori_per_anno (senza
    interrompere l'elaborazione degli altri anni).
    """
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=False, keep_links=True)
        workbook_valori = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=True)
    except InvalidFileException as exc:
        raise ExcelTemplateError(
            "Il file template non e' un file Excel (.xlsx) valido."
        ) from exc
    except Exception as exc:
        raise ExcelTemplateError(
            f"Impossibile aprire il file Excel template: {exc}"
        ) from exc

    sheet = _get_sheet(workbook, SHEET_NAME_RICLASSIFICAZIONE)
    sheet_valori = _get_sheet(workbook_valori, SHEET_NAME_RICLASSIFICAZIONE)

    colonne_anno = _rileva_colonne_anno(sheet_valori)
    mappa_anno_colonna = {anno_val: col for col, anno_val in colonne_anno.items()}

    report_per_anno: Dict[str, WriteReport] = {}
    errori_per_anno: Dict[str, str] = {}

    for anno, risultato in risultati_per_anno.items():
        colonna = mappa_anno_colonna.get(str(anno))
        if colonna is None:
            anni_trovati = ", ".join(sorted(mappa_anno_colonna.keys())) or "nessuno"
            errori_per_anno[anno] = (
                f"Il template non contiene una colonna per l'anno '{anno}'. "
                f"Anni individuati automaticamente nel template: {anni_trovati}."
            )
            continue
        report_per_anno[anno] = _scrivi_valori_anno(sheet, colonna, risultato)

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    return buffer, report_per_anno, errori_per_anno


# ---------------------------------------------------------------------------
# FUNZIONE PUBBLICA: SCRITTURA A SINGOLO ANNO (compatibilita')
# ---------------------------------------------------------------------------
def popola_template_excel(template_file, anno: str, risultato: RiclassificazioneResult):
    """Variante a singolo anno, mantenuta per compatibilita': internamente
    usa lo stesso motore multi-anno con un solo elemento."""
    dati = _leggi_bytes(template_file)
    buffer, report_per_anno, errori_per_anno = popola_template_multi(dati, {anno: risultato})

    if anno in errori_per_anno:
        raise ExcelMappingError(errori_per_anno[anno])

    return buffer, report_per_anno[anno]
