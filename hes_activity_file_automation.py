"""
Automate HES Activity workbook processing.

This script:
- compares headers
- renames worksheet tabs
- archives old workbooks
- promotes new workbooks

"""
# Version 1.2
#
# Completed:
# - workbook discovery
# - header validation
# - worksheet renaming
# - archive process
# - promote incoming workbook
# - stop on header mismatch
# - main() workflow
#
# Outstanding:
# - modified date warning
# - production testing

import pandas as pd

from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime
# shutil moves files
import shutil

DATA_FOLDER = Path(
    r"X:\Energy Assistance Package\Users\Joanna.Pogorzelska-Abbasciano"
    r"\Project\Product analysis"
    r"\PowerBI reports\Data sources"
)

PRODUCTION_FILE = DATA_FOLDER / "HES Activity.xlsx"

ARCHIVE_FOLDER = DATA_FOLDER / "Archive"

#print(DATA_FOLDER)
#print(PRODUCTION_FILE)
#print(ARCHIVE_FOLDER)

def find_new_file():
    
    """
    Find incoming Excel workbooks
    excluding the production workbook.
    """
    
    excel_files = [
        file
        for file in DATA_FOLDER.glob("*.xlsx")
        if file.name != "HES Activity.xlsx"
     ] 

    if len(excel_files) !=1:
        raise ValueError(
            f"Expected 1 incoming workbook, "
            f"found {len(excel_files)}"
        )

    return excel_files[0]




def get_headers(file_path):
    """
    Return the column headers from Excel workbook.
    """

    headers = list(
        pd.read_excel(
            file_path,
            nrows=0
        ).columns
    )

    return headers

def headers_match(file_one, file_two):
    """
    Check whether both workbooks have identical column headers.
    """

    headers_one = get_headers(file_one)
    headers_two = get_headers(file_two)

    return headers_one == headers_two

def validate_headers(file_path):
    """
    Validate that both files have identical headers.
    """
    # progress message
    print("Comparing column headers...")

    if headers_match(
        PRODUCTION_FILE,
        file_path
    ):
        print("PASS: Headers match")
        return True
    
    raise ValueError(
        "Headers do not match"
    )

def get_sheet_name(file_path):
    """
    Return the name of the first worksheet.
    """

    workbook = load_workbook(file_path)

    return workbook.active.title

def rename_worksheet(file_path):
    """
    Rename to worksheet to 'HES Activity'
    """
    
    # progress message
    print("Opening workbook...")

    workbook = load_workbook(file_path)
    
    # progress message
    print("Renaming worksheet tab...")

    workbook.active.title = "HES Activity"

    # progress message
    print("Saving workbook...")


    workbook.save(file_path)

    # progress message
    print("Worksheet renamed succesfully...")

def archive_old_file():
    """
    Archive the current production workbook.
    """
    print("Creating archive folder...")

    ARCHIVE_FOLDER.mkdir(exist_ok=True)

    archive_date = datetime.now().strftime("%Y%m%d")

    archive_file = (
        ARCHIVE_FOLDER
        / f"HES Activity_archived_{archive_date}.xlsx"
    )

    # progress message
    print("Archiving production workbook...")

    shutil.move(
        PRODUCTION_FILE,
        archive_file
    )

    # progress message
    print("Production workbook archived.")

def promote_new_file(file_path):
    """
    Promote the incoming workbook to the production workbook.
    """
    print("Promoting incoming workbook...")

    file_path.rename(PRODUCTION_FILE)

    print("Incoming workbook promoted.")

def check_file_date(file_path):
    """
    Warn if the incoming workbook was not modified today.
    """
    # file_path.stat() - tell me everything about that file
    # st_mtime - last modified time ; .date() - gets rid of HH:MM:SS
    file_modified_date = datetime.fromtimestamp(
        file_path.stat().st_mtime
    ).date()

    today = datetime.now().date()

    if file_modified_date == today:
        print("PASS: Workbook was modified today.")

    else:
        print(
            "WARNING: Incoming workbook: "
            f"was last modified on "
            f"{file_modified_date}"
        )



def main():
    """
    Execute the complete HES workbook replacement workflow.
    """
    
    # progress message
    print("Finding incoming workbook...")

    new_file = find_new_file()

    check_file_date(new_file)

    validate_headers(new_file)

    rename_worksheet(new_file)

    archive_old_file()

    promote_new_file(new_file)

    # progress message
    print("Checking worksheet name...")

    print(
        f"Worksheet name is now: " 
        f"{get_sheet_name(PRODUCTION_FILE)}"
    )
# __name__ is a special Python variable.
# It equals "__main__" only when this file is run directly.
# This prevents the workflow from running automatically
# if the file is imported by another script.
if __name__ == "__main__":
    main()