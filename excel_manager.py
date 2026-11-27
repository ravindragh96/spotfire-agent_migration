
import openpyxl
from pathlib import Path
from datetime import datetime

from config import EXCEL_FILE


# ==========================================
# MODULE 2: EXCEL MANAGER
# ==========================================

class ExcelManager:

    def __init__(self, excel_file=EXCEL_FILE):

        self.excel_file = Path(excel_file)
        self.sheet_name = "Migration_Requests"

        self.required_columns = [
            "Migration_ID",
            "Report_Name",
            "Source_Path",
            "Migration_Type",
            "Destination_Path",
            "Status",
            "Backup_Path",
            "Last_Updated",
            "Error_Message"
        ]

    # --------------------------------------
    # 1. Check Excel file
    # --------------------------------------

    def check_file(self):

        if not self.excel_file.exists():
            raise FileNotFoundError(
                f"Excel tracker not found: {self.excel_file}"
            )

        print("Excel tracker found.")

    # --------------------------------------
    # 2. Load workbook and worksheet
    # --------------------------------------

    def load_workbook(self):

        self.check_file()

        workbook = openpyxl.load_workbook(
            self.excel_file
        )

        if self.sheet_name not in workbook.sheetnames:
            workbook.close()
            raise ValueError(
                f"Worksheet '{self.sheet_name}' not found."
            )

        worksheet = workbook[self.sheet_name]

        return workbook, worksheet

    # --------------------------------------
    # 3. Validate Excel column headers
    # --------------------------------------

    def validate_headers(self, worksheet):

        actual_headers = [
            cell.value
            for cell in worksheet[1]
        ]

        actual_headers = actual_headers[:len(self.required_columns)]

        if actual_headers != self.required_columns:
            raise ValueError(
                "Excel headers do not match the required format.\n"
                f"Expected: {self.required_columns}\n"
                f"Found: {actual_headers}"
            )

        return True

    # --------------------------------------
    # 4. Read all migration requests
    # --------------------------------------

    def get_all_requests(self):

        workbook, worksheet = self.load_workbook()

        try:
            self.validate_headers(worksheet)

            requests = []

            for row in worksheet.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not any(value is not None for value in row):
                    continue

                request = dict(
                    zip(self.required_columns, row)
                )

                requests.append(request)

            return requests

        finally:
            workbook.close()

    # --------------------------------------
    # 5. Check duplicate migration IDs
    # --------------------------------------

    def validate_unique_ids(self, requests):

        migration_ids = [
            str(request["Migration_ID"]).strip()
            for request in requests
            if request["Migration_ID"] is not None
        ]

        if len(migration_ids) != len(set(migration_ids)):
            raise ValueError(
                "Duplicate Migration_ID found in Excel."
            )

        return True

    # --------------------------------------
    # 6. Get pending migration requests
    # --------------------------------------

    def get_pending_requests(self):

        requests = self.get_all_requests()

        self.validate_unique_ids(requests)

        pending_requests = []

        for request in requests:

            status = str(
                request.get("Status") or ""
            ).strip().lower()

            if status == "pending":
                pending_requests.append(request)

        return pending_requests

    # --------------------------------------
    # 7. Find request using Migration ID
    # --------------------------------------

    def get_request_by_id(self, migration_id):

        requests = self.get_all_requests()

        for request in requests:

            current_id = str(
                request.get("Migration_ID") or ""
            ).strip()

            if current_id == str(migration_id).strip():
                return request

        return None

    # --------------------------------------
    # 8. Update migration status
    # --------------------------------------

    def update_request(
        self,
        migration_id,
        status=None,
        backup_path=None,
        error_message=None
    ):

        workbook, worksheet = self.load_workbook()

        try:
            self.validate_headers(worksheet)

            matching_row = None

            for row_number in range(2, worksheet.max_row + 1):

                current_id = worksheet.cell(
                    row=row_number,
                    column=1
                ).value

                if str(current_id or "").strip() == str(migration_id).strip():

                    if matching_row is not None:
                        raise ValueError(
                            f"Duplicate Migration_ID: {migration_id}"
                        )

                    matching_row = row_number

            if matching_row is None:
                raise ValueError(
                    f"Migration ID not found: {migration_id}"
                )

            # Column F: Status
            if status is not None:

                allowed_statuses = [
                    "Pending",
                    "In Progress",
                    "Completed",
                    "Failed",
                    "Needs Review"
                ]

                if status not in allowed_statuses:
                    raise ValueError(
                        f"Invalid status: {status}"
                    )

                worksheet.cell(
                    row=matching_row,
                    column=6
                ).value = status

            # Column G: Backup Path
            if backup_path is not None:

                worksheet.cell(
                    row=matching_row,
                    column=7
                ).value = str(backup_path)

            # Column H: Last Updated
            worksheet.cell(
                row=matching_row,
                column=8
            ).value = datetime.now().astimezone().isoformat(
                timespec="seconds"
            )

            # Column I: Error Message
            if error_message is not None:

                worksheet.cell(
                    row=matching_row,
                    column=9
                ).value = error_message

            workbook.save(self.excel_file)

            print(
                f"Migration {migration_id} updated successfully."
            )

        finally:
            workbook.close()

    # --------------------------------------
    # 9. Display migration requests
    # --------------------------------------

    def display_requests(self, requests):

        if not requests:
            print("No migration requests found.")
            return

        for request in requests:

            print("\n------------------------------")

            for key, value in request.items():

                print(f"{key}: {value}")

        print("\n------------------------------")


# ==========================================
# MODULE TEST
# ==========================================

if __name__ == "__main__":

    excel_manager = ExcelManager()

    print("\n===== ALL MIGRATION REQUESTS =====")

    all_requests = excel_manager.get_all_requests()

    excel_manager.validate_unique_ids(all_requests)
    excel_manager.display_requests(all_requests)

    print("\n===== PENDING MIGRATION REQUESTS =====")

    pending_requests = excel_manager.get_pending_requests()

    excel_manager.display_requests(pending_requests)
