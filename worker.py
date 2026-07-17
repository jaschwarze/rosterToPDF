"""
Enthält die eigentliche Verarbeitungslogik (vormals in main.py), jetzt
parametrisiert und als QThread, damit die GUI während der PDF-Erzeugung
nicht einfriert.
"""

import glob
import os
import shutil
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QThread, Signal

from pdf import create_employee_view, create_group_view, create_leader_view
from parser import parse_employee_times

DAYS_OF_WEEK = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]


class GenerationError(Exception):
    """Fehler, die während der Verarbeitung auftreten und dem Nutzer
    verständlich angezeigt werden sollen."""


class PdfGenerationWorker(QThread):
    log = Signal(str)
    progress = Signal(int, int)  # (aktueller Schritt, Schritte insgesamt)
    finished_ok = Signal(str)    # Erfolgsmeldung
    finished_error = Signal(str)  # Fehlermeldung

    def __init__(self, excel_path: str, output_path: str, archive_path: str,
                 cols_per_day: int = 6, parent=None):
        super().__init__(parent)
        self.excel_path = excel_path
        self.output_path = output_path
        self.archive_path = archive_path
        self.cols_per_day = cols_per_day

    def run(self):
        try:
            self._generate()
        except GenerationError as exc:
            self.finished_error.emit(str(exc))
        except Exception as exc:  # unerwarteter Fehler
            self.finished_error.emit(f"Unerwarteter Fehler: {exc}")

    def _generate(self):
        excel_file_path = self.excel_path
        output_path = self.output_path
        archive_path = self.archive_path
        cols_per_day = self.cols_per_day

        if not Path(excel_file_path).exists():
            raise GenerationError("Die ausgewählte Excel-Datei existiert nicht mehr.")

        for path in [Path(output_path), Path(archive_path)]:
            if not path.exists():
                path.mkdir(parents=True)

        self.log.emit("Lese Excel-Datei ein...")
        employee_data = pd.read_excel(
            excel_file_path, sheet_name="Mitarbeiterliste",
            skiprows=2, header=None, usecols="A:C, E:G"
        )
        special_dates_data = pd.read_excel(
            excel_file_path, sheet_name="Sondertermine", skiprows=2, header=None
        )
        planning_data = pd.read_excel(
            excel_file_path, sheet_name="Dienstplanung", header=None
        )

        employee_dict = {
            row[0]: (row[1], row[2])
            for row in employee_data.itertuples(index=False)
        }

        special_dates_dict = {
            row[0]: (row[1], row[2], row[3], row[4], row[5])
            for row in special_dates_data.itertuples(index=True)
        }

        possible_assignments = {}
        for row in employee_data.itertuples(index=False):
            if pd.notna(row[3]):
                assignment = row[3]
                abbreviation = row[4] if pd.notna(row[4]) else ""
                color_code = row[5] if pd.notna(row[5]) else ""

                possible_assignments[assignment] = {
                    "abbreviation": abbreviation,
                    "color": color_code
                }

        possible_groups = list(possible_assignments.keys())[:6]

        year = planning_data[1][0]
        calendar_week = planning_data[1][1]
        start_date = planning_data[1][3].strftime("%d.%m.%Y")
        end_date = planning_data[1][5].strftime("%d.%m.%Y")

        planning_frame = planning_data.iloc[12:]
        employee_times = parse_employee_times(planning_frame, cols_per_day, DAYS_OF_WEEK)

        self.log.emit("Erstelle Mitarbeiteransicht... (1/3)")
        self.progress.emit(1, 4)
        create_employee_view(
            employee_times, output_path, possible_assignments, year,
            calendar_week, start_date, DAYS_OF_WEEK, special_dates_dict
        )

        self.log.emit("Erstelle Gruppenansicht... (2/3)")
        self.progress.emit(2, 4)
        create_group_view(
            employee_times, output_path, possible_assignments, year,
            calendar_week, start_date, DAYS_OF_WEEK, possible_groups,
            employee_dict, special_dates_dict
        )

        self.log.emit("Erstelle Leitungsansicht... (3/3)")
        self.progress.emit(3, 4)
        create_leader_view(
            employee_times, output_path, possible_assignments, year,
            calendar_week, DAYS_OF_WEEK, possible_groups, employee_dict
        )

        copy_path = os.path.join(archive_path, str(year), "KW-" + str(calendar_week))
        if os.path.isdir(copy_path):
            self.log.emit(f"Kopie der Auswertung in {copy_path} übersprungen, da sie schon existiert.")
        else:
            os.makedirs(copy_path, exist_ok=True)
            for file in glob.glob(os.path.join(output_path, "*.pdf")):
                output_file = os.path.join(copy_path, os.path.basename(file))
                shutil.copyfile(file, output_file)
            self.log.emit(f"Archivkopie erstellt unter {copy_path}.")

        self.progress.emit(4, 4)
        self.finished_ok.emit(
            f"Fertig! Pläne für KW {calendar_week}/{year} wurden erstellt "
            f"({start_date} - {end_date})."
        )
