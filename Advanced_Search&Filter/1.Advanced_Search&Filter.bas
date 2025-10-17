Attribute VB_Name = "Module1"
'===========================================================
' Project : Excel VBA Automation
' Module  : Advanced_Search_Filter.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 1.0
'===========================================================
'
' Description:
' Filters the dataset on "DataSheet" (A:E) by optional
' criteria typed in:
'   G2 = Department
'   G3 = Region
'   G4 = Status
' and writes the result to a new sheet "FilteredResults".
' - Blank criteria are ignored
' - Case-insensitive, partial match (contains)
' - Copies header and visible rows, keeps column widths
'
' Usage:
'   - Attach a button to Run_AdvancedFilter
'   - (Optional) Attach another button to Clear_FilterResults
'
' -----------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -----------------------------------------------------------
' © 2025 Data Solutions Lab. by Osman Uluhan
'===========================================================

Option Explicit

Public Sub Run_AdvancedFilter()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim rngData As Range, rngVisible As Range
    Dim lastRow As Long
    Dim critDept As String, critRegion As String, critStatus As String
    Dim hadResults As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '--- base sheet & data range
    Set ws = ThisWorkbook.Worksheets("DataSheet")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found on DataSheet.", vbExclamation
        GoTo SafeExit
    End If
    Set rngData = ws.Range("A1:E" & lastRow)
    
    '--- read criteria (empty = ignore)
    critDept = Trim$(ws.Range("H2").Value)
    critRegion = Trim$(ws.Range("H3").Value)
    critStatus = Trim$(ws.Range("H4").Value)
    
    '--- clear any prior filter
    If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData
    
    '--- apply AutoFilter (contains, case-insensitive)
    rngData.AutoFilter
    If Len(critDept) > 0 Then rngData.AutoFilter Field:=1, Criteria1:="=*" & critDept & "*"
    If Len(critRegion) > 0 Then rngData.AutoFilter Field:=3, Criteria1:="=*" & critRegion & "*"
    If Len(critStatus) > 0 Then rngData.AutoFilter Field:=5, Criteria1:="=*" & critStatus & "*"
    
    '--- get visible rows
    On Error Resume Next
    Set rngVisible = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandler
    
    '--- (re)create output sheet
    DeleteSheetIfExists "FilteredResults"
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ws)
    wsOut.Name = "FilteredResults"
    
    '--- copy header
    rngData.Rows(1).Copy Destination:=wsOut.Range("A1")
    
    '--- copy filtered rows (exclude header if no data)
    hadResults = False
    Dim area As Range
    Dim destRow As Long
    destRow = 2

    If Not rngVisible Is Nothing Then
        For Each area In rngVisible.Areas
            ' Skip header row (A1:E1)
            If area.Row > 1 Then
                area.Copy Destination:=wsOut.Range("A" & destRow)
                destRow = destRow + area.Rows.Count
                hadResults = True
            End If
        Next area
    End If
    
    '--- format widths to match source
    rngData.Rows(1).EntireColumn.Copy
    wsOut.Range("A1").EntireColumn.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    
    '--- friendly message when nothing found
    If Not hadResults Then
        MsgBox "No rows matched the given criteria.", vbInformation
    Else
        MsgBox "Filtering completed successfully.", vbInformation
    End If
    
SafeExit:
    ' clear filter on data sheet (keep original view clean)
    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.AutoFilter.ShowAllData
        On Error GoTo 0
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit
End Sub


Public Sub Clear_FilterResults()
    ' Deletes the result sheet if exists and shows all data
    On Error Resume Next
    DeleteSheetIfExists "FilteredResults"
    On Error GoTo 0
    
    With ThisWorkbook.Worksheets("DataSheet")
        If .AutoFilterMode Then
            On Error Resume Next
            .AutoFilter.ShowAllData
            On Error GoTo 0
        End If
    End With
    
    MsgBox "Results cleared.", vbInformation
End Sub


'-----------------------
' helpers
'-----------------------
Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Worksheets
        If LCase$(sh.Name) = LCase$(sheetName) Then
            sh.Delete
            Exit For
        End If
    Next sh
    Application.DisplayAlerts = True
End Sub
