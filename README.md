# About
Just a simple Powershell script to export calendar items from Outlook to a CSV file and afterwards updates the Excel sheet where the CSV is imported. No rights reserved.

# Export your Outlook Calendar Items to a CSV file

```
powershell .\ExportOutlookCalenderToCSV.ps1 -StartDate "01-01-2019" -EndDate "31-12-2019" -ExportCsvLocation "export_calendar_2019.csv" -ExcelLocation "registratie_2019.xlsx" -Format '#,##'
```