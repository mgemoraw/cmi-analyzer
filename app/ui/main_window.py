# app/gui/main_window.py

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QFileDialog,
    QProgressBar,
    QGroupBox,
    QSpinBox,
    QMessageBox,
    QFrame,
    QSizePolicy,
)

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("Excel Batch Processor")

        self.resize(1400, 900)

        self.setMinimumSize(1200, 800)

        self.build_ui()

    # =========================================================
    # BUILD UI
    # =========================================================

    def build_ui(self):

        # CENTRAL WIDGET
        central_widget = QWidget()

        self.setCentralWidget(central_widget)

        # MAIN LAYOUT
        self.main_layout = QVBoxLayout()

        self.main_layout.setContentsMargins(20, 20, 20, 20)

        self.main_layout.setSpacing(15)

        central_widget.setLayout(self.main_layout)

        # HEADER
        self.build_header()

        # SETTINGS SECTION
        self.build_settings_section()

        # ACTION BUTTONS
        self.build_action_buttons()

        # PROGRESS SECTION
        self.build_progress_section()

        # STATISTICS SECTION
        self.build_statistics_section()

        # LOG PANEL
        self.build_log_panel()

        # STATUS BAR
        self.statusBar().showMessage("Ready")

    # =========================================================
    # HEADER
    # =========================================================

    def build_header(self):

        container = QFrame()

        layout = QVBoxLayout()

        container.setLayout(layout)

        title = QLabel("Excel Batch Processor")

        title.setFont(QFont("Segoe UI", 24, QFont.Bold))

        subtitle = QLabel(
            "Professional Excel Processing System"
        )

        subtitle.setFont(QFont("Segoe UI", 11))

        subtitle.setStyleSheet(
            "color: #aaaaaa;"
        )

        layout.addWidget(title)

        layout.addWidget(subtitle)

        self.main_layout.addWidget(container)

    # =========================================================
    # SETTINGS SECTION
    # =========================================================

    def build_settings_section(self):

        group = QGroupBox("Input Settings")

        group.setObjectName("settingsGroup")

        layout = QGridLayout()

        layout.setVerticalSpacing(15)

        layout.setHorizontalSpacing(10)

        group.setLayout(layout)

        # INPUT FOLDER
        input_label = QLabel("Input Folder")

        self.input_edit = QLineEdit()

        self.input_edit.setPlaceholderText(
            "Select input folder..."
        )

        input_btn = QPushButton("Browse")

        input_btn.clicked.connect(
            self.select_input_folder
        )

        layout.addWidget(input_label, 0, 0)

        layout.addWidget(self.input_edit, 0, 1)

        layout.addWidget(input_btn, 0, 2)

        # OUTPUT FOLDER
        output_label = QLabel("Output Folder")

        self.output_edit = QLineEdit()

        self.output_edit.setPlaceholderText(
            "Select output folder..."
        )

        output_btn = QPushButton("Browse")

        output_btn.clicked.connect(
            self.select_output_folder
        )

        layout.addWidget(output_label, 1, 0)

        layout.addWidget(self.output_edit, 1, 1)

        layout.addWidget(output_btn, 1, 2)

        # CONFIG FILE
        config_label = QLabel("Config File")

        self.config_edit = QLineEdit()

        self.config_edit.setPlaceholderText(
            "Select config.yaml..."
        )

        config_btn = QPushButton("Browse")

        config_btn.clicked.connect(
            self.select_config_file
        )

        layout.addWidget(config_label, 2, 0)

        layout.addWidget(self.config_edit, 2, 1)

        layout.addWidget(config_btn, 2, 2)

        # CHUNK SIZE
        chunk_label = QLabel("Chunk Size")

        self.chunk_spin = QSpinBox()

        self.chunk_spin.setMinimum(1)

        self.chunk_spin.setMaximum(100000)

        self.chunk_spin.setValue(100)

        layout.addWidget(chunk_label, 3, 0)

        layout.addWidget(self.chunk_spin, 3, 1)

        # WORKERS
        worker_label = QLabel("Workers")

        self.worker_spin = QSpinBox()

        self.worker_spin.setMinimum(1)

        self.worker_spin.setMaximum(64)

        self.worker_spin.setValue(4)

        layout.addWidget(worker_label, 4, 0)

        layout.addWidget(self.worker_spin, 4, 1)

        self.main_layout.addWidget(group)

    # =========================================================
    # ACTION BUTTONS
    # =========================================================

    def build_action_buttons(self):

        container = QFrame()

        layout = QHBoxLayout()

        layout.setSpacing(10)

        container.setLayout(layout)

        # START
        self.start_btn = QPushButton(
            "Start Processing"
        )

        self.start_btn.setObjectName("startButton")

        self.start_btn.setMinimumHeight(45)

        # PAUSE
        self.pause_btn = QPushButton("Pause")

        self.pause_btn.setMinimumHeight(45)

        # RESUME
        self.resume_btn = QPushButton("Resume")

        self.resume_btn.setMinimumHeight(45)

        # STOP
        self.stop_btn = QPushButton("Stop")

        self.stop_btn.setMinimumHeight(45)

        # OPEN OUTPUT
        self.open_output_btn = QPushButton(
            "Open Output Folder"
        )

        self.open_output_btn.setMinimumHeight(45)

        layout.addWidget(self.start_btn)

        layout.addWidget(self.pause_btn)

        layout.addWidget(self.resume_btn)

        layout.addWidget(self.stop_btn)

        layout.addStretch()

        layout.addWidget(self.open_output_btn)

        self.main_layout.addWidget(container)

    # =========================================================
    # PROGRESS SECTION
    # =========================================================

    def build_progress_section(self):

        group = QGroupBox("Processing Status")

        layout = QVBoxLayout()

        group.setLayout(layout)

        # CURRENT FILE
        self.current_file_label = QLabel(
            "Current File: None"
        )

        layout.addWidget(self.current_file_label)

        # PROGRESS BAR
        self.progress_bar = QProgressBar()

        self.progress_bar.setValue(0)

        self.progress_bar.setMinimumHeight(30)

        layout.addWidget(self.progress_bar)

        # PERCENTAGE
        self.progress_percentage = QLabel("0%")

        self.progress_percentage.setAlignment(
            Qt.AlignRight
        )

        layout.addWidget(self.progress_percentage)

        self.main_layout.addWidget(group)

    # =========================================================
    # STATISTICS SECTION
    # =========================================================

    def build_statistics_section(self):

        group = QGroupBox("Statistics")

        layout = QGridLayout()

        layout.setVerticalSpacing(15)

        group.setLayout(layout)

        # FILES
        files_title = QLabel("Files Processed")

        self.files_value = QLabel("0")

        # EQUIPMENT
        eq_title = QLabel("Equipment Rows")

        self.eq_value = QLabel("0")

        # MPDM
        mpdm_title = QLabel("MPDM Rows")

        self.mpdm_value = QLabel("0")

        # DAILY
        daily_title = QLabel("DailyVariables Rows")

        self.daily_value = QLabel("0")

        # ERRORS
        error_title = QLabel("Errors")

        self.error_value = QLabel("0")

        stats_widgets = [
            (files_title, self.files_value),
            (eq_title, self.eq_value),
            (mpdm_title, self.mpdm_value),
            (daily_title, self.daily_value),
            (error_title, self.error_value),
        ]

        row = 0

        for title, value in stats_widgets:

            title.setFont(
                QFont("Segoe UI", 10)
            )

            value.setFont(
                QFont("Segoe UI", 16, QFont.Bold)
            )

            value.setAlignment(Qt.AlignCenter)

            value.setMinimumHeight(50)

            value.setObjectName("statValue")

            layout.addWidget(title, row, 0)

            layout.addWidget(value, row, 1)

            row += 1

        self.main_layout.addWidget(group)

    # =========================================================
    # LOG PANEL
    # =========================================================

    def build_log_panel(self):

        group = QGroupBox("Live Logs")

        layout = QVBoxLayout()

        group.setLayout(layout)

        self.log_text = QTextEdit()

        self.log_text.setReadOnly(True)

        self.log_text.setPlaceholderText(
            "Logs will appear here..."
        )

        self.log_text.setSizePolicy(
            QSizePolicy.Expanding,
            QSizePolicy.Expanding
        )

        layout.addWidget(self.log_text)

        self.main_layout.addWidget(group)

    # =========================================================
    # FILE SELECTORS
    # =========================================================

    def select_input_folder(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Input Folder"
        )

        if folder:
            self.input_edit.setText(folder)

    def select_output_folder(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder"
        )

        if folder:
            self.output_edit.setText(folder)

    def select_config_file(self):

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Config File",
            filter="YAML Files (*.yaml)"
        )

        if file_path:
            self.config_edit.setText(file_path)

    # =========================================================
    # LOGGING
    # =========================================================

    def add_log(self, message):

        self.log_text.append(message)

    # =========================================================
    # PROGRESS UPDATE
    # =========================================================

    def update_progress(self, value):

        self.progress_bar.setValue(value)

        self.progress_percentage.setText(
            f"{value}%"
        )

    # =========================================================
    # UPDATE STATISTICS
    # =========================================================

    def update_statistics(
        self,
        files=0,
        equipment=0,
        mpdm=0,
        daily=0,
        errors=0
    ):

        self.files_value.setText(str(files))

        self.eq_value.setText(str(equipment))

        self.mpdm_value.setText(str(mpdm))

        self.daily_value.setText(str(daily))

        self.error_value.setText(str(errors))

    # =========================================================
    # ERROR DIALOG
    # =========================================================

    def show_error(self, message):

        QMessageBox.critical(
            self,
            "Error",
            message
        )

    # =========================================================
    # SUCCESS DIALOG
    # =========================================================

    def show_success(self, message):

        QMessageBox.information(
            self,
            "Success",
            message
        )



# IMPORTANT ADDITIONS FOR main_window.py

# ==========================================
# ADD THESE IMPORTS
# ==========================================

from PySide6.QtCore import QThread

from workers.processing_worker import (
    ProcessingWorker
)

from core.utils import Utils

# ==========================================
# ADD THIS INSIDE __init__()
# ==========================================

self.thread = None

self.worker = None

# ==========================================
# CONNECT BUTTONS
# ==========================================

self.start_btn.clicked.connect(
    self.start_processing
)

self.open_output_btn.clicked.connect(
    self.open_output_folder
)

# ==========================================
# ADD THESE METHODS
# ==========================================

def start_processing(self):

    input_dir = self.input_edit.text().strip()

    output_dir = self.output_edit.text().strip()

    if not input_dir:

        self.show_error(
            "Please select input folder."
        )

        return

    if not output_dir:

        self.show_error(
            "Please select output folder."
        )

        return

    self.progress_bar.setValue(0)

    self.log_text.clear()

    self.add_log(
        "[INFO] Starting processing..."
    )

    self.thread = QThread()

    self.worker = ProcessingWorker(

        input_dir=input_dir,

        output_dir=output_dir,

        chunk_size=self.chunk_spin.value(),

        max_workers=self.worker_spin.value()
    )

    self.worker.moveToThread(
        self.thread
    )

    # ======================================
    # SIGNAL CONNECTIONS
    # ======================================

    self.thread.started.connect(
        self.worker.run
    )

    self.worker.finished.connect(
        self.thread.quit
    )

    self.worker.finished.connect(
        self.worker.deleteLater
    )

    self.thread.finished.connect(
        self.thread.deleteLater
    )

    self.worker.log.connect(
        self.add_log
    )

    self.worker.progress.connect(
        self.update_progress
    )

    self.worker.statistics.connect(
        self.handle_statistics
    )

    self.worker.current_file.connect(
        self.update_current_file
    )

    self.worker.error.connect(
        self.show_error
    )

    self.worker.finished.connect(
        self.processing_finished
    )

    self.start_btn.setEnabled(False)

    self.thread.start()

def processing_finished(self):

    self.add_log(
        "[SUCCESS] Processing completed."
    )

    self.show_success(
        "Excel processing completed."
    )

    self.start_btn.setEnabled(True)

def update_current_file(self, filename):

    self.current_file_label.setText(
        f"Current File: {filename}"
    )

def handle_statistics(self, stats):

    self.update_statistics(

        files=stats["files"],

        equipment=stats["equipment"],

        mpdm=stats["mpdm"],

        daily=stats["daily"],

        errors=stats["errors"]
    )

def open_output_folder(self):

    output_dir = self.output_edit.text().strip()

    if output_dir:

        Utils.open_folder(output_dir)