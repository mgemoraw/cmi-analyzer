# app/core/validator.py

import pandas as pd


class Validator:
    """
    Validation Engine
    """

    # ==========================================
    # INIT
    # ==========================================

    def __init__(
        self,
        required_sheets=None,
        required_headers=None,
        logger_callback=None
    ):

        self.required_sheets = (
            required_sheets or []
        )

        self.required_headers = (
            required_headers or {}
        )

        self.logger_callback = (
            logger_callback
        )

    # ==========================================
    # LOGGER
    # ==========================================

    def log(self, message):

        print(message)

        if self.logger_callback:
            self.logger_callback(message)

    # ==========================================
    # VALIDATE SHEETS
    # ==========================================

    def validate_sheets(
        self,
        available_sheets
    ):

        missing_sheets = []

        normalized = [
            s.lower()
            for s in available_sheets
        ]

        for sheet in self.required_sheets:

            if sheet.lower() not in normalized:

                missing_sheets.append(sheet)

        return missing_sheets

    # ==========================================
    # VALIDATE HEADERS
    # ==========================================

    def validate_headers(
        self,
        dataframe,
        sheet_name
    ):

        missing_headers = []

        required = self.required_headers.get(
            sheet_name,
            []
        )

        columns = [
            str(col).strip()
            for col in dataframe.columns
        ]

        for header in required:

            if header not in columns:

                missing_headers.append(header)

        return missing_headers

    # ==========================================
    # REMOVE EMPTY ROWS
    # ==========================================

    def remove_empty_rows(
        self,
        dataframe
    ):

        if dataframe.empty:
            return dataframe

        dataframe = dataframe.dropna(
            how="all"
        )

        dataframe = dataframe.reset_index(
            drop=True
        )

        return dataframe

    # ==========================================
    # REMOVE DUPLICATES
    # ==========================================

    def remove_duplicates(
        self,
        dataframe
    ):

        if dataframe.empty:
            return dataframe

        dataframe = dataframe.drop_duplicates()

        dataframe = dataframe.reset_index(
            drop=True
        )

        return dataframe

    # ==========================================
    # CLEAN DATAFRAME
    # ==========================================

    def clean_dataframe(
        self,
        dataframe
    ):

        dataframe = self.remove_empty_rows(
            dataframe
        )

        dataframe = self.remove_duplicates(
            dataframe
        )

        return dataframe

    # ==========================================
    # VALIDATE DATAFRAME
    # ==========================================

    def validate_dataframe(
        self,
        dataframe,
        sheet_name
    ):

        errors = []

        if dataframe.empty:

            errors.append(
                f"{sheet_name}: Empty dataframe"
            )

            return errors

        missing_headers = (
            self.validate_headers(
                dataframe,
                sheet_name
            )
        )

        if missing_headers:

            errors.append(
                f"{sheet_name}: Missing headers -> "
                f"{missing_headers}"
            )

        return errors