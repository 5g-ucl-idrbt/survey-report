# surveyreport

Steps to Automate Using Excel and VBA

Open Your Excel File:

Open the "Survey Responses.xlsx" file.

**add a row below the Administrative section**

Create a New Module for VBA:

Press ALT + F11 to open the VBA editor.

In the VBA editor, go to Insert > Module to create a new module.

Write the VBA Code:

Copy and paste the following VBA code into the new module:

```
Sub CompileSurveyResponsesHorizontally()
    Dim ws As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsSheet1 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim col As Long

    ' Set worksheets
    Set wsTemplate = ThisWorkbook.Sheets("Question Template")
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")

    ' Clear Sheet1
    wsSheet1.Cells.Clear

    ' Copy questions to Sheet1
    wsTemplate.Range("A1:A" & wsTemplate.Cells(wsTemplate.Rows.Count, "A").End(xlUp).Row).Copy Destination:=wsSheet1.Range("A1")
    wsTemplate.Range("B1:B" & wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row).Copy Destination:=wsSheet1.Range("B1")

    ' Set initial column for responses
    col = 3

    ' Loop through each sheet except "Question Template" and "Sheet1"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Question Template" And ws.Name <> "Sheet1" Then
            ' Add respondent name as header
            wsSheet1.Cells(1, col).Value = ws.Name

            ' Copy responses to Sheet1
            lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
            ws.Range("C1:C" & lastRow).Copy
            wsSheet1.Cells(2, col).PasteSpecial Paste:=xlPasteValues

            ' Move to the next column
            col = col + 1
        End If
    Next ws

    ' Autofit columns
    wsSheet1.Columns.AutoFit
End Sub

```
Run the VBA Code:

Close the VBA editor.
Press ALT + F8, select CompileSurveyResponses, and click Run.

select All copy then create new sheet right click on first cell select paste special > paste special > Transpose

This VBA script will:

Clear the contents of "Sheet1".

Copy the questions from the "Question Template" sheet to "Sheet1".

Loop through each individual's sheet and copy their responses to "Sheet1".

Label each column in "Sheet1" with the corresponding respondent's name.






