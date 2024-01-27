# Attendance Tracker

## Overview
The Attendance Tracker is designed to process and track employee attendance data for IKEA PL Activity Hours. This project automates the process of filtering, copying, processing, and tracking attendance data, ensuring efficient and accurate management of employee attendance records.

## Workflow
The workflow involves several steps, from obtaining raw data in an Excel file to processing it and updating an attendance tracker. The process is automated using Python scripts and involves updating a Git repository.

### Prerequisites
- Excel File named "IKEA PL Activity Hours" provided by Hyun Wook Jung, Intraday Specialist.
- Python installed with necessary libraries for running scripts.
- Git and Git Bash installed for version control.
- Access to the Codespaces repository.

### Steps

1. **Prepare the Excel File**
   - Obtain the "IKEA PL Activity Hours" Excel file.
   - Open the file and filter for Occurrences (Unexcused Absence, LOA, NCL, Partial, Tardy).

2. **Process Raw Data**
   - Copy the entire filtered workbook to `RawData.txt`.

3. **Run Scripts in Codespaces**
   - Access the repository in Codespaces: [AttendanceTracker Codespaces](https://github.com/codespaces?repository_id=749111140)
   - Execute `RawDataProcess.py` and verify the output in `summaries.txt`.
   - Execute `AttendanceProcess.py` and verify the output in `AttendanceTracker.xlsx`.
   - Commit changes in CodeSpace.

4. **Format Attendance Tracker**
   - Open `AttendanceTracker.xlsx` located at `C:\Users\OLMAL21\OneDrive - IKEA\Desktop\AttendanceTracker`.
   - Format and colorize the attendance tracker as needed.

5. **Git Version Control**
   - Open Git Bash.
   - Navigate to the AttendanceTracker directory:
     ```
     cd "/c/Users/OLMAL21/OneDrive - IKEA/Desktop/AttendanceTracker"
     ```
   - Add the updated `AttendanceTracker.xlsx` to Git:
     ```
     git add AttendanceTracker.xlsx
     ```
   - Commit the changes with a meaningful message:
     ```
     git commit -m "Updated Attendance Tracker for [Date/Description]"
     ```
   - Push the changes to the main branch:
     ```
     git push origin main
     ```
   - Pull any updates from the main branch:
     ```
     git pull origin main
     ```

## Notes
- Ensure you have the latest version of the scripts and Excel file before beginning the process.
- Always verify the outputs after running Python scripts to ensure data integrity.
- Keep your local and remote repositories synchronized by regularly pulling and pushing changes.
