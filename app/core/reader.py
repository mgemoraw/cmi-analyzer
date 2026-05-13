# app/core/excel_reader.py

from pathlib import Path

import pandas as pd


class ExcelReader:
    """
    Excel Reader Class

    Responsibilities:
    -----------------
    1. Read Excel workbook
    2. Detect required sheets
    3. Clean empty rows
    4. Return structured data
    """

    REQUIRED_SHEETS = [
        "equipment",
        "mpdm",
        "dailyvariables"
    ]

    # =====================================================
    # INIT
    # =====================================================

    def __init__(self, logger_callback=None):

        self.logger_callback = logger_callback

    # =====================================================
    # LOGGER
    # =====================================================

    def log(self, message):

        print(message)

        if self.logger_callback:
            self.logger_callback(message)

    # =====================================================
    # FIND SHEET NAME
    # =====================================================

    def find_sheet_name(
        self,
        available_sheets,
        target_sheet
    ):

        sheet_map = {
            sheet.lower(): sheet
            for sheet in available_sheets
        }

        return sheet_map.get(
            target_sheet.lower()
        )

    # =====================================================
    # CLEAN DATAFRAME
    # =====================================================

    def clean_dataframe(self, df):

        # REMOVE FULLY EMPTY ROWS
        df = df.dropna(how="all")

        # RESET INDEX
        df = df.reset_index(drop=True)

        return df

    # =====================================================
    # READ SINGLE SHEET
    # =====================================================

    def read_sheet(
        self,
        file_path,
        sheet_name
    ):

        try:

            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                engine="openpyxl"
            )

            df = self.clean_dataframe(df)

            return df

        except Exception as e:

            self.log(
                f"[ERROR] Failed reading "
                f"sheet '{sheet_name}': {str(e)}"
            )

            return pd.DataFrame()

    # =====================================================
    # READ WORKBOOK
    # =====================================================

    def read_workbook(self, file_path):

        file_path = Path(file_path)

        self.log(
            f"[INFO] Reading workbook: "
            f"{file_path.name}"
        )

        result = {
            "equipment": pd.DataFrame(),
            "mpdm": pd.DataFrame(),
            "dailyvariables": pd.DataFrame(),
        }

        try:

            excel = pd.ExcelFile(file_path)

            available_sheets = excel.sheet_names

            # ==============================
            # EQUIPMENT
            # ==============================

            equipment_sheet = self.find_sheet_name(
                available_sheets,
                "equipment"
            )

            if equipment_sheet:

                result["equipment"] = (
                    self.read_sheet(
                        file_path,
                        equipment_sheet
                    )
                )

            else:

                self.log(
                    f"[WARNING] equipment sheet "
                    f"missing in {file_path.name}"
                )

            # ==============================
            # MPDM
            # ==============================

            mpdm_sheet = self.find_sheet_name(
                available_sheets,
                "mpdm"
            )

            if mpdm_sheet:

                result["mpdm"] = (
                    self.read_sheet(
                        file_path,
                        mpdm_sheet
                    )
                )

            else:

                self.log(
                    f"[WARNING] mpdm sheet "
                    f"missing in {file_path.name}"
                )

            # ==============================
            # DAILYVARIABLES
            # ==============================

            daily_sheet = self.find_sheet_name(
                available_sheets,
                "dailyvariables"
            )

            if daily_sheet:

                result["dailyvariables"] = (
                    self.read_sheet(
                        file_path,
                        daily_sheet
                    )
                )

            else:

                self.log(
                    f"[WARNING] dailyvariables "
                    f"sheet missing in "
                    f"{file_path.name}"
                )

            return result

        except Exception as e:

            self.log(
                f"[ERROR] Failed reading workbook "
                f"{file_path.name}: {str(e)}"
            )

            return result