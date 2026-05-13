"""
Excel Batch Splitter
====================

This script:

1. Reads Excel files from an input folder
2. Extracts data from:
    - equipment sheet
    - mpdm sheet
    - dailyvariables sheet

3. Combines ALL dailyvariables records into ONE Excel file

4. Splits equipment and mpdm records into chunks of 100 rows
   and saves them into separate Excel files.

Designed with Object-Oriented Programming (OOP)
using pandas + openpyxl.

------------------------------------------------
INSTALLATION
------------------------------------------------

pip install pandas openpyxl

------------------------------------------------
FOLDER STRUCTURE
------------------------------------------------

project/
│
├── input_files/
│     file1.xlsx
│     file2.xlsx
│     ...
│
├── output/
│
└── excel_splitter.py

------------------------------------------------
EXPECTED SHEET NAMES
------------------------------------------------

equipment
mpdm
dailyvariables

(case-insensitive)

------------------------------------------------
HOW TO RUN
------------------------------------------------

python excel_splitter.py
"""

from pathlib import Path
import pandas as pd
from math import ceil


class ExcelReader:
    """
    Reads Excel files and extracts required sheets.
    """

    REQUIRED_SHEETS = ["equipment", "mpdm", "dailyvariables"]

    def __init__(self, file_path):
        self.file_path = Path(file_path)

    def read_sheet(self, sheet_name):
        """
        Read a sheet safely.
        """
        try:
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                engine="openpyxl"
            )

            # Remove completely empty rows
            df = df.dropna(how="all")

            return df

        except Exception as e:
            print(f"Error reading {sheet_name} in {self.file_path.name}: {e}")
            return pd.DataFrame()

    def read_all_required_sheets(self):
        """
        Returns dictionary of all required sheets.
        """
        excel_file = pd.ExcelFile(self.file_path)

        # Make sheet matching case-insensitive
        available_sheets = {
            s.lower(): s for s in excel_file.sheet_names
        }

        result = {}

        for sheet in self.REQUIRED_SHEETS:
            if sheet.lower() in available_sheets:
                actual_name = available_sheets[sheet.lower()]
                result[sheet] = self.read_sheet(actual_name)
            else:
                print(f"Missing sheet '{sheet}' in {self.file_path.name}")
                result[sheet] = pd.DataFrame()

        return result


class DataCollector:
    """
    Collects data from multiple Excel files.
    """

    def __init__(self):
        self.equipment_data = []
        self.mpdm_data = []
        self.dailyvariables_data = []

    def add_data(self, data_dict):
        """
        Add dataframes into memory lists.
        """

        if not data_dict["equipment"].empty:
            self.equipment_data.append(data_dict["equipment"])

        if not data_dict["mpdm"].empty:
            self.mpdm_data.append(data_dict["mpdm"])

        if not data_dict["dailyvariables"].empty:
            self.dailyvariables_data.append(data_dict["dailyvariables"])

    def get_combined_equipment(self):
        return pd.concat(self.equipment_data, ignore_index=True)

    def get_combined_mpdm(self):
        return pd.concat(self.mpdm_data, ignore_index=True)

    def get_combined_dailyvariables(self):
        return pd.concat(self.dailyvariables_data, ignore_index=True)


class ExcelWriter:
    """
    Writes Excel output files.
    """

    def __init__(self, output_dir="output"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)

    def save_dataframe(self, df, file_name, sheet_name="Sheet1"):
        """
        Save dataframe into Excel.
        """
        output_path = self.output_dir / file_name

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)

        print(f"Saved: {output_path}")

    def split_and_save(self, df, prefix, chunk_size=100):
        """
        Split dataframe into chunks and save each chunk.
        """

        total_rows = len(df)

        if total_rows == 0:
            print(f"No data found for {prefix}")
            return

        total_files = ceil(total_rows / chunk_size)

        for i in range(total_files):
            start = i * chunk_size
            end = start + chunk_size

            chunk_df = df.iloc[start:end]

            file_name = f"{prefix}_{i + 1}.xlsx"

            self.save_dataframe(
                chunk_df,
                file_name=file_name,
                sheet_name=prefix
            )


class ExcelBatchProcessor:
    """
    Main application controller.
    """

    def __init__(self, input_folder="input_files", output_folder="output"):
        self.input_folder = Path(input_folder)
        self.writer = ExcelWriter(output_folder)
        self.collector = DataCollector()

    def process_files(self):
        """
        Read all Excel files.
        """

        excel_files = list(self.input_folder.glob("*.xlsx"))

        if not excel_files:
            print("No Excel files found.")
            return

        print(f"Found {len(excel_files)} Excel files")

        for file in excel_files:
            print(f"Processing: {file.name}")

            reader = ExcelReader(file)
            data = reader.read_all_required_sheets()

            self.collector.add_data(data)

        self.generate_outputs()

    def generate_outputs(self):
        """
        Generate final output files.
        """

        # Combine all data
        equipment_df = self.collector.get_combined_equipment()
        mpdm_df = self.collector.get_combined_mpdm()
        daily_df = self.collector.get_combined_dailyvariables()

        # Save ALL dailyvariables into one file
        self.writer.save_dataframe(
            daily_df,
            "all_dailyvariables.xlsx",
            sheet_name="dailyvariables"
        )

        # Split equipment into 100-row files
        self.writer.split_and_save(
            equipment_df,
            prefix="equipment",
            chunk_size=100
        )

        # Split mpdm into 100-row files
        self.writer.split_and_save(
            mpdm_df,
            prefix="mpdm",
            chunk_size=100
        )

        print("\nProcessing completed successfully!")


if __name__ == "__main__":

    processor = ExcelBatchProcessor(
        input_folder="input_files",
        output_folder="output"
    )

    processor.process_files()