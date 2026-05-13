# app/core/logger.py

from pathlib import Path
from datetime import datetime


class AppLogger:

    # ==========================================
    # INIT
    # ==========================================

    def __init__(self, log_dir="logs"):

        self.log_dir = Path(log_dir)

        self.log_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        self.log_file = (
            self.log_dir
            / f"log_{timestamp}.txt"
        )

    # ==========================================
    # WRITE LOG
    # ==========================================

    def write(self, message):

        timestamp = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        log_message = (
            f"[{timestamp}] {message}"
        )

        with open(
            self.log_file,
            "a",
            encoding="utf-8"
        ) as file:

            file.write(log_message + "\n")