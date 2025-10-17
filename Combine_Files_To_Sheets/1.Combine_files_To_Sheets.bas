Attribute VB_Name = "Module1"

' ============================================================
'  Project : Excel VBA Automation
'  Module  : Combine_Files_To_Sheets.bas
'  Author  : Osman Uluhan
'  Date    : 2025-10-13
'  Version : 2.0
' ------------------------------------------------------------
'  Description:
'  This VBA macro combines multiple Excel files into a single
'  workbook — creating a separate worksheet for each source
'  file, while keeping original formatting and column widths.
'
'  Features:
'   - Prompts user to select a folder with Excel files
'   - Copies all sheets (with formatting) into one workbook
'   - Preserves fonts, borders, colors, and cell width
'   - Automatically renames sheets based on file names
'   - Works with both .xlsx and .xlsm files
'
'  License:
'  MIT License – You are free to use, modify, and distribute
'  this code with attribution.
'
'  © 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
' ============================================================

Sub Combine_Files_To_Sheets()
    Dim FolderPath As String
    Dim FileName As String
    Dim wbSource As Workbook
    Dim wbMaster As Workbook
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim SheetName As String
    
    ' --- Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Excel Files"
        If .Show <> -1 Then Exit Sub
        FolderPath = .SelectedItems(1) & "\"
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- Create new workbook for merged result
    Set wbMaster = Workbooks.Add
    
    ' --- Loop through all Excel files
    FileName = Dir(FolderPath & "*.xls*")
    Do While FileName <> ""
        Set wbSource = Workbooks.Open(FolderPath & FileName)
        
        ' Take only the first sheet (or loop if multi-sheet)
        For Each wsSource In wbSource.Worksheets
            SheetName = Left(wsSource.Name, 25)
            On Error Resume Next
            Set wsNew = wbMaster.Sheets(SheetName)
            If wsNew Is Nothing Then
                wsSource.Copy After:=wbMaster.Sheets(wbMaster.Sheets.Count)
                wbMaster.Sheets(wbMaster.Sheets.Count).Name = _
                    Left(Replace(FileName, ".xlsx", ""), 28)
            End If
            On Error GoTo 0
        Next wsSource
        
        wbSource.Close SaveChanges:=False
        FileName = Dir
    Loop
    
    ' --- Remove default blank sheet if still exists
    On Error Resume Next
    Application.DisplayAlerts = False
    wbMaster.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    MsgBox "? All Excel files have been combined successfully!" & vbCrLf & _
           "Each file was imported as a separate sheet with formatting preserved.", _
           vbInformation, "Process Completed"
End Sub


