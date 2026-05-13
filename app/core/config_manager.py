# app/core/config_manager.py

from pathlib import Path

import yaml


class ConfigManager:
    """
    Configuration Loader
    """

    # ==========================================
    # INIT
    # ==========================================

    def __init__(self, config_path):

        self.config_path = Path(config_path)

        self.config_data = {}

    # ==========================================
    # LOAD CONFIG
    # ==========================================

    def load(self):

        if not self.config_path.exists():

            raise FileNotFoundError(
                f"Config file not found: "
                f"{self.config_path}"
            )

        with open(
            self.config_path,
            "r",
            encoding="utf-8"
        ) as file:

            self.config_data = yaml.safe_load(file)

        return self.config_data

    # ==========================================
    # GET VALUE
    # ==========================================

    def get(
        self,
        key,
        default=None
    ):

        return self.config_data.get(
            key,
            default
        )

    # ==========================================
    # GET NESTED VALUE
    # ==========================================

    def get_nested(
        self,
        *keys,
        default=None
    ):

        data = self.config_data

        try:

            for key in keys:

                data = data[key]

            return data

        except Exception:

            return default