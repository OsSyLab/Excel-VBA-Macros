Attribute VB_Name = "Module1"
Sub Automatic_Report_Generator()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long, reportRow As Long
    Dim sheetName As String

    Application.ScreenUpdating = False

    ' Create or clear "Report" sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Report")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "Report"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0

    reportRow = 1

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsReport.Name Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Copy headers from the first sheet only
            If reportRow = 1 Then
                ws.Rows(1).Copy wsReport.Rows(1)
                wsReport.Cells(1, wsReport.Cells(1, Columns.Count).End(xlToLeft).Column + 1).Value = "Sheet Name"
                reportRow = 2
            End If

            ' Copy data rows
            ws.Range("A2:A" & lastRow).EntireRow.Copy wsReport.Cells(reportRow, 1)

            ' Add sheet name in the last column
            sheetName = ws.Name
            wsReport.Range(wsReport.Cells(reportRow, wsReport.Cells(1, Columns.Count).End(xlToLeft).Column), _
                           wsReport.Cells(wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row, wsReport.Cells(1, Columns.Count).End(xlToLeft).Column)).Value = sheetName

            ' Update next row
            reportRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row + 1
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "? Report has been successfully generated!", vbInformation
End Sub


