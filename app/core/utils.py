# app/core/utils.py

import os
import subprocess
import platform


class Utils:

    # ==========================================
    # OPEN FOLDER
    # ==========================================

    @staticmethod
    def open_folder(path):

        if not os.path.exists(path):
            return

        system = platform.system()

        if system == "Windows":

            os.startfile(path)

        elif system == "Darwin":

            subprocess.Popen(["open", path])

        else:

            subprocess.Popen(["xdg-open", path])