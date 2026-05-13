# app/workers/processing_worker.py

from PySide6.QtCore import QObject
from PySide6.QtCore import Signal
from PySide6.QtCore import Slot

from core.excel_processor import ExcelProcessor


class ProcessingWorker(QObject):

    # ==========================================
    # SIGNALS
    # ==========================================

    finished = Signal()

    error = Signal(str)

    progress = Signal(int)

    log = Signal(str)

    statistics = Signal(dict)

    current_file = Signal(str)

    # ==========================================
    # INIT
    # ==========================================

    def __init__(
        self,
        input_dir,
        output_dir,
        chunk_size,
        max_workers
    ):

        super().__init__()

        self.input_dir = input_dir

        self.output_dir = output_dir

        self.chunk_size = chunk_size

        self.max_workers = max_workers

    # ==========================================
    # RUN
    # ==========================================

    @Slot()
    def run(self):

        try:

            processor = ExcelProcessor(

                input_dir=self.input_dir,

                output_dir=self.output_dir,

                chunk_size=self.chunk_size,

                max_workers=self.max_workers,

                logger_callback=self.emit_log,

                progress_callback=self.emit_progress,

                statistics_callback=self.emit_statistics,

                current_file_callback=self.emit_current_file
            )

            processor.process()

            self.finished.emit()

        except Exception as e:

            self.error.emit(str(e))

    # ==========================================
    # EMIT LOG
    # ==========================================

    def emit_log(self, message):

        self.log.emit(message)

    # ==========================================
    # EMIT PROGRESS
    # ==========================================

    def emit_progress(self, value):

        self.progress.emit(value)

    # ==========================================
    # EMIT CURRENT FILE
    # ==========================================

    def emit_current_file(self, filename):

        self.current_file.emit(filename)

    # ==========================================
    # EMIT STATISTICS
    # ==========================================

    def emit_statistics(
        self,
        files,
        equipment,
        mpdm,
        daily,
        errors
    ):

        stats = {

            "files": files,

            "equipment": equipment,

            "mpdm": mpdm,

            "daily": daily,

            "errors": errors,
        }

        self.statistics.emit(stats)