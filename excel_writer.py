# -*- coding: utf-8 -*-
"""
excel_writer.py
===============
Scrittura dei dati riclassificati nel file Excel "Template" tramite
openpyxl, senza alterare formule o formattazioni preesistenti.

openpyxl con `keep_vba=False, data_only=False` preserva formule e stili
del workbook originale: scriviamo solo i valori nelle celle di destinazione,
lasciando intatto tutto il resto (intestazioni, formule di calcolo margine,
colori, bordi, ecc.).
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from typing import List

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from config import CELL_MAPPING, SHEET_NAME_RICLASSIFICAZIONE
from copilot_client import RiclassificazioneResult


class ExcelTemplateError(Exception):
    """Errore nell'apertura o nella struttura del file Excel template."""


class ExcelMappingError(Exception):
    """Errore nel mapping tra categorie e celle per l'anno selezionato."""


@dataclass
class WriteReport:
    """Riepilogo di ciò che è stato scritto, utile per un feedback chiaro in UI."""
    celle_scritte: int
    voci_non_mappate: List[str]
    foglio: str


def _get_sheet(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    # fallback: usa il foglio attivo se il nome configurato non esiste,
    # ma segnala comunque un avviso al chiamante.
    return workbook.active


def popola_template_excel(template_file, anno: str, risultato: RiclassificazioneResult) -> tuple[io.BytesIO, WriteReport]:
    """
    Apre il template Excel caricato dall'utente, scrive i valori
    riclassificati nelle celle corrispondenti all'anno selezionato e
    restituisce un buffer in memoria pronto per il download, insieme a un
    report di cosa è stato effettivamente scritto.
    """
    if anno not in CELL_MAPPING:
        raise ExcelMappingError(
            f"Non esiste un mapping di celle configurato per l'anno '{anno}'. "
            "Aggiorna CELL_MAPPING in config.py per includere questo anno."
        )

    try:
        template_file.seek(0)
        workbook = openpyxl.load_workbook(template_file, data_only=False, keep_links=True)
    except InvalidFileException as exc:
        raise ExcelTemplateError(
            "Il file caricato non è un file Excel (.xlsx) valido."
        ) from exc
    except Exception as exc:
        raise ExcelTemplateError(
            f"Impossibile aprire il file Excel template: {exc}"
        ) from exc

    sheet = _get_sheet(workbook, SHEET_NAME_RICLASSIFICAZIONE)
    mapping_anno = CELL_MAPPING[anno]

    valori_da_scrivere = {
        **risultato.entrate,
        **risultato.uscite,
        **risultato.gestione_fiscale,
    }

    celle_scritte = 0
    voci_non_mappate: List[str] = []

    for categoria, valore in valori_da_scrivere.items():
        cella = mapping_anno.get(categoria)
        if cella is None:
            voci_non_mappate.append(categoria)
            continue
        try:
            sheet[cella] = valore
            celle_scritte += 1
        except Exception as exc:
            raise ExcelTemplateError(
                f"Errore scrivendo la categoria '{categoria}' nella cella '{cella}': {exc}"
            ) from exc

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    report = WriteReport(
        celle_scritte=celle_scritte,
        voci_non_mappate=voci_non_mappate,
        foglio=sheet.title,
    )
    return buffer, report
