from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QLabel, QPushButton, QFileDialog, QPlainTextEdit, QProgressBar,
    QMessageBox, QSizePolicy
)

from settings_manager import SettingsManager
from worker import PdfGenerationWorker


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = SettingsManager()
        self.excel_path: str | None = None
        self.worker: PdfGenerationWorker | None = None

        self.setWindowTitle("Dienstplanerstellung")
        self.resize(1020, 820)

        self._build_menu()
        self._build_central_widget()
        self._load_saved_folders()
        self._update_start_button_state()

    # ------------------------------------------------------------------
    # Aufbau der Oberfläche
    # ------------------------------------------------------------------
    def _build_menu(self):
        file_menu = self.menuBar().addMenu("&Datei")

        open_action = QAction("Excel-Datei öffnen...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._choose_excel_file)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        exit_action = QAction("Beenden", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def _build_central_widget(self):
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # --- Eingabedatei ---
        input_box = QGroupBox("Eingabedatei")
        input_layout = QHBoxLayout(input_box)
        self.excel_label = QLabel("Keine Datei ausgewählt")
        self.excel_label.setStyleSheet("color: #b00020;")
        self.excel_label.setWordWrap(True)
        choose_excel_btn = QPushButton("Excel-Datei wählen...")
        choose_excel_btn.clicked.connect(self._choose_excel_file)
        input_layout.addWidget(self.excel_label, stretch=1)
        input_layout.addWidget(choose_excel_btn)
        layout.addWidget(input_box)

        # --- Ordner ---
        folder_box = QGroupBox("Ordner")
        folder_layout = QVBoxLayout(folder_box)

        self.output_label = QLabel()
        self.output_label.setWordWrap(True)
        output_row = self._build_folder_row(
            "Ausgangsordner:", self.output_label, self._choose_output_folder
        )
        folder_layout.addLayout(output_row)

        self.archive_label = QLabel()
        self.archive_label.setWordWrap(True)
        archive_row = self._build_folder_row(
            "Archivordner:", self.archive_label, self._choose_archive_folder
        )
        folder_layout.addLayout(archive_row)

        layout.addWidget(folder_box)

        # --- Start ---
        self.start_button = QPushButton("PDF-Erzeugung starten")
        self.start_button.setMinimumHeight(40)
        self.start_button.clicked.connect(self._start_generation)
        layout.addWidget(self.start_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 4)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # --- Log ---
        log_box = QGroupBox("Verlauf")
        log_layout = QVBoxLayout(log_box)
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        log_layout.addWidget(self.log_output)
        layout.addWidget(log_box, stretch=1)

        self.setCentralWidget(central)

    @staticmethod
    def _build_folder_row(label_text: str, path_label: QLabel, on_choose) -> QHBoxLayout:
        row = QHBoxLayout()
        title = QLabel(label_text)
        title.setMinimumWidth(120)
        path_label.setText("(noch nicht ausgewählt)")
        button = QPushButton("Ändern...")
        button.clicked.connect(on_choose)
        row.addWidget(title)
        row.addWidget(path_label, stretch=1)
        row.addWidget(button)
        return row

    # ------------------------------------------------------------------
    # Einstellungen laden
    # ------------------------------------------------------------------
    def _load_saved_folders(self):
        if self.settings.output_path:
            self.output_label.setText(self.settings.output_path)
        if self.settings.archive_path:
            self.archive_label.setText(self.settings.archive_path)

    # ------------------------------------------------------------------
    # Aktionen
    # ------------------------------------------------------------------
    def _choose_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Excel-Datei auswählen", "", "Excel-Dateien (*.xlsx)"
        )
        if path:
            self.excel_path = path
            self.excel_label.setText(path)
            self.excel_label.setStyleSheet("color: #1b5e20;")
        self._update_start_button_state()

    def _choose_output_folder(self):
        start_dir = self.settings.output_path or ""
        folder = QFileDialog.getExistingDirectory(self, "Ausgangsordner wählen", start_dir)
        if folder:
            self.settings.output_path = folder
            self.output_label.setText(folder)
        self._update_start_button_state()

    def _choose_archive_folder(self):
        start_dir = self.settings.archive_path or ""
        folder = QFileDialog.getExistingDirectory(self, "Archivordner wählen", start_dir)
        if folder:
            self.settings.archive_path = folder
            self.archive_label.setText(folder)
        self._update_start_button_state()

    def _update_start_button_state(self):
        ready = bool(
            self.excel_path
            and self.settings.output_path
            and self.settings.archive_path
        )
        self.start_button.setEnabled(ready)

    def _start_generation(self):
        if not self.excel_path:
            QMessageBox.warning(self, "Keine Datei", "Bitte zuerst eine Excel-Datei auswählen.")
            return

        self.start_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_output.clear()
        self._log("Starte PDF-Erzeugung...")

        self.worker = PdfGenerationWorker(
            excel_path=self.excel_path,
            output_path=self.settings.output_path,
            archive_path=self.settings.archive_path,
            cols_per_day=self.settings.cols_per_day,
        )
        self.worker.log.connect(self._log)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_ok.connect(self._on_finished_ok)
        self.worker.finished_error.connect(self._on_finished_error)
        self.worker.start()

    # ------------------------------------------------------------------
    # Callbacks des Worker-Threads
    # ------------------------------------------------------------------
    def _log(self, message: str):
        self.log_output.appendPlainText(message)

    def _on_progress(self, step: int, total: int):
        self.progress_bar.setRange(0, total)
        self.progress_bar.setValue(step)

    def _on_finished_ok(self, message: str):
        self._log(message)
        QMessageBox.information(self, "Fertig", message)
        self._update_start_button_state()

    def _on_finished_error(self, message: str):
        self._log(f"FEHLER: {message}")
        QMessageBox.critical(self, "Fehler", message)
        self._update_start_button_state()
