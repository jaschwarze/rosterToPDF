"""
Kümmert sich um das Speichern/Laden der Einstellungen, die zwischen
Programmstarts erhalten bleiben sollen (Ausgangs- und Archivordner).

Die Auswahl der Excel-Datei wird bewusst NICHT gespeichert, da sie bei
jedem Programmstart neu getroffen werden soll.

QSettings legt die Daten plattformspezifisch ab (Windows: Registry,
macOS: plist, Linux: ini-Datei) - es ist also kein eigenes
Konfigurationsformat nötig und funktioniert auch aus der .exe heraus.
"""

from PySide6.QtCore import QSettings

ORG_NAME = "Dienstplan-Tool"
APP_NAME = "Dienstplanerstellung"

KEY_OUTPUT_PATH = "paths/output_path"
KEY_ARCHIVE_PATH = "paths/archive_path"
KEY_COLS_PER_DAY = "options/cols_per_day"


class SettingsManager:
    def __init__(self):
        self._settings = QSettings(ORG_NAME, APP_NAME)

    @property
    def output_path(self) -> str:
        return self._settings.value(KEY_OUTPUT_PATH, "", type=str)

    @output_path.setter
    def output_path(self, value: str) -> None:
        self._settings.setValue(KEY_OUTPUT_PATH, value)

    @property
    def archive_path(self) -> str:
        return self._settings.value(KEY_ARCHIVE_PATH, "", type=str)

    @archive_path.setter
    def archive_path(self, value: str) -> None:
        self._settings.setValue(KEY_ARCHIVE_PATH, value)

    @property
    def cols_per_day(self) -> int:
        return self._settings.value(KEY_COLS_PER_DAY, 6, type=int)

    @cols_per_day.setter
    def cols_per_day(self, value: int) -> None:
        self._settings.setValue(KEY_COLS_PER_DAY, value)
