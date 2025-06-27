# PS Scripts Collection
Automates daily reporting to reduce errors and streamline repetitive tasks.

## Prerequisites
- IText7Module

## Scripts Included
- `ExtractMSGFiles for Daily Report (run once).ps1` - Extracts all attachments from MSG files. These attachments contain the reports.

- `#Master.ps1` - Runs all the scripts in order
- `VeeamReportPDF.ps1` - Identifies errors from the last 6 days and organises the file by date.
- `WeeklyChangeReportPDF.ps1` - Organises the file by date.
- `WeeklyServiceReportPDF.ps1` - Organises the file by date.
- `WsusMSG.ps1` - Organises the file by date.
- `ActifioReportPDF.ps1` - Identifies errors organises the file by date.
- `ConnectBackupsMSG.ps1` - Identifies errors organises the file by date.
- `IPSReportPDF.ps1` - Identifies errors organises the file by date.
- `iSeriesPDF.ps1` - Identifies errors organises the file by date.
- `SymantecHTML.ps1` - Organises the file by date.
- `TrendAVReportPDF.ps1` - Organises the file by date.

- `CheckForReportWithError.ps1` - Checks if any files have an error at the end of the name and alerts the user to review the report.

## How to Use
1.  **Prepare the Reports Folder**
    * Ensure all daily report emails are stored in the `reports` folder in MSG format.

2.  **Run the First Script**
    * Execute the `ExtractMSGFiles for Daily Report (run once).ps1` script.
    * Run it only once to avoid extracting the PDFs/HTMLs reports multiple times.

3.  **Run the Master Script**
    * Execute the `Master.ps1` script.
    * This script will loop through all the daily report scripts.
    * It will only check the `reports` folder for the reports.
    * Ensure the reports are in the correct format (e.g., PDF, MSG, HTML).

4.  **Handle Errors**
    * The last script in the loop is `CheckForReportWithError.ps1`.
    * If any of the scripts find an error in a file/report, they will rename the file with "error" at the end.
    * The `CheckForReportWithError.ps1` script will alert you when it encounters a file with "error" at the end.

5.  **Stop the Master Script**
    * The `Master.ps1` script will keep looping every minute until you manually close it.