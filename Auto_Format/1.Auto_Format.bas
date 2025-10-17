Attribute VB_Name = "Module1"
' Project : Excel VBA Automation
' Module  : Auto_Format_On_Open.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 2.0
' =======================================================
'
' Description:
' Automatically applies consistent formatting when the workbook opens.
' Formats font, header row, borders, alignment, and resets mixed styles.
'
' Works on "Sheet1" and ensures a clean, professional Excel layout.
'
' -------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -------------------------------------------------------
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =======================================================

Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim lastInRows As Range, lastInCols As Range
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range, hdr As Range
    Dim defaultRowHeight As Double

    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- SAFE last used cell detection (no type mismatch) ----
    ' Search by rows
    Set lastInRows = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    ' Search by columns
    Set lastInCols = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False)

    ' Empty sheet? exit gracefully
    If lastInRows Is Nothing Or lastInCols Is Nothing Then GoTo SafeExit

    lastRow = lastInRows.Row
    lastCol = lastInCols.Column

    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Set hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))   ' only filled header cells

    ' ---- Reset & unify fonts/styles for ALL used cells ----
    With rng
        .Interior.Color = RGB(255, 255, 255)            ' white body
        With .Font
            .Name = "Calibri"
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Color = RGB(0, 0, 0)
        End With
    End With

    ' ---- Header style (only A1:lastCol in row 1) ----
    With hdr
        .Interior.Color = RGB(217, 225, 242)            ' light blue
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With

    ' ---- Borders on the used range ----
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(180, 180, 180)
        .Weight = xlThin
    End With

    ' ---- Alignment ----
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter

    ' ---- Column widths & equal row heights ----
    rng.Columns.AutoFit
    defaultRowHeight = 20
    ws.Rows("1:" & lastRow).RowHeight = defaultRowHeight   ' every row same height

    ' Optional confirmation
    MsgBox "? Workbook formatted." & vbCrLf & _
           "Header: A1 to " & ws.Cells(1, lastCol).Address(False, False) & vbCrLf & _
           "Rows set to equal height: " & defaultRowHeight & " pt.", _
           vbInformation, "Auto Format Ready!"

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
