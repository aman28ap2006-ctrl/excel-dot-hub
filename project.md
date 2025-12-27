Excel Automation uses VBA (Visual Basic for Applications) to:
Automate repetitive tasks
Clean & process data
Generate reports
Import/export files
Save time and reduce errors
##Dataset used
= <a href="https://github.com/aman28ap2006-ctrl/excel-dot-hub/blob/main/AUTOMATION.xlsx">Dataset: viwe </a>
Excel-Automation-Using-VBA/
│
├── README.md
├── macros/
│   ├── data_cleaning.bas
│   ├── report_generator.bas
│   └── automation.bas
│
├── sample_files/
│   └── sample_data.xlsx
│
└── screenshots/
    └── output.png
    🔹 Data Cleaning Macro
    - dashboard <a href="https://github.com/aman28ap2006-ctrl/excel-dot-hub/blob/main/screensort%20of%20automation.xlsx">dashboard view</a>
Sub DataCleaning()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    <img width="2195" height="1172" alt="image" src="https://github.com/user-attachments/assets/39ac5671-49d7-4ed2-baf6-270e844e6512" />
   'Remove blank rows
    ws.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp

    'Remove duplicates from column A
    ws.Range("A1:D1000").RemoveDuplicates Columns:=1, Header:=xlYes

    MsgBox "Data Cleaning Completed Successfully!", vbInformation
End Sub
Sub GenerateReport()
    Dim reportSheet As Worksheet
    Set reportSheet = Sheets.Add
    reportSheet.Name = "Summary Report"

    reportSheet.Range("A1").Value = "Total Records"
    reportSheet.Range("B1").Value = WorksheetFunction.CountA(Sheets(1).Range("A:A"))

    MsgBox "Report Generated!", vbInformation
End Sub
Sub AutoSaveFile()
    Dim path As String
    path = ThisWorkbook.Path & "\Report_" & Format(Date, "dd-mm-yyyy") & ".xlsx"
    ThisWorkbook.SaveAs path
End Sub
README.md 
# Excel Automation Using VBA

## Features
- Automated Data Cleaning
- Report Generation
- File Auto-Save
- Error Reduction

## Tools Used
- Microsoft Excel
- VBA Macros
- GitHub

## How to Use
1. Enable Macros in Excel
2. Import `.bas` files
3. Run macros from Developer Tab
