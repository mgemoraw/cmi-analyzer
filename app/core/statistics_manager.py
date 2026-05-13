# app/core/statistics_manager.py


class StatisticsManager:
    """
    Statistics Manager

    Tracks:
    -------
    - processed files
    - equipment rows
    - mpdm rows
    - dailyvariables rows
    - errors
    """

    # ==========================================
    # INIT
    # ==========================================

    def __init__(self):

        self.reset()

    # ==========================================
    # RESET
    # ==========================================

    def reset(self):

        self.files_processed = 0

        self.equipment_rows = 0

        self.mpdm_rows = 0

        self.daily_rows = 0

        self.errors = 0

    # ==========================================
    # UPDATE
    # ==========================================

    def update(
        self,
        equipment_count=0,
        mpdm_count=0,
        daily_count=0,
        errors=0
    ):

        self.files_processed += 1

        self.equipment_rows += (
            equipment_count
        )

        self.mpdm_rows += (
            mpdm_count
        )

        self.daily_rows += (
            daily_count
        )

        self.errors += errors

    # ==========================================
    # TO DICTIONARY
    # ==========================================

    def to_dict(self):

        return {

            "files": self.files_processed,

            "equipment": self.equipment_rows,

            "mpdm": self.mpdm_rows,

            "daily": self.daily_rows,

            "errors": self.errors,
        }