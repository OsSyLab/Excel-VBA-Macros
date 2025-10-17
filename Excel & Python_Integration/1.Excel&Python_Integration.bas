Attribute VB_Name = "Module1"
'========================================
'  Excel + Python Integration Macro
'  Exports Excel data › runs Python script › imports result
'  Created by Osman Uluhan
'========================================

' ==============================================
' Excel + Python Integration (Robust v2)
' - Saves CSV to a TEMP folder with unique name
' - Passes paths to Python as arguments
' - Waits until result file exists (timeout safe)
' Author : Osman Uluhan
' ==============================================

Sub RunPythonIntegration()
    Dim ws As Worksheet
    Dim tempDir As String, ts As String
    Dim dataPath As String, resultPath As String
    Dim pyExe As String, pyScript As String
    Dim cmd As String, timeout As Single, t As Single
    Dim lastRow As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets("DataSheet")  ' Target worksheet

    ' === Temporary folder ===
    tempDir = Environ$("TEMP") & "\xl_py\"
    If Dir(tempDir, vbDirectory) = "" Then MkDir tempDir

    ts = Format(Now, "yyyymmdd_hhnnss")
    dataPath = tempDir & "data_" & ts & ".csv"
    resultPath = tempDir & "result_" & ts & ".csv"
    MsgBox "Output CSV Path:" & vbCrLf & resultPath

    ' === Export range A1:B to CSV ===
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("A1:B" & lastRow).Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:=dataPath, FileFormat:=xlCSV, Local:=True
    ActiveWorkbook.Close SaveChanges:=False

    ' === Python paths ===
    pyExe = "C:\Users\Osman ULUHAN\AppData\Local\Programs\Python\Python313\python.exe"
    pyScript = "C:\PythonExcelTest\script.py"

    cmd = Chr(34) & pyExe & Chr(34) & " " & _
          Chr(34) & pyScript & Chr(34) & " " & _
          Chr(34) & dataPath & Chr(34) & " " & _
          Chr(34) & resultPath & Chr(34)

    MsgBox cmd
    Shell cmd, vbNormalFocus

    ' === Wait for Python output ===
    timeout = 10
    t = Timer
    Do While Dir(resultPath) = ""
        DoEvents
        If Timer - t > timeout Then
            MsgBox "? Timeout: Python output not found.", vbCritical
            Exit Sub
        End If
    Loop

    ' === Clear previous results ===
    ws.Range("C1:D" & ws.Rows.Count).ClearContents

    ' === Import the result CSV ===
    With ws.QueryTables.Add(Connection:="TEXT;" & resultPath, Destination:=ws.Range("C1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With

    ' === Copy formatting from columns A & B to C & D ===
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim lastResultRow As Long
    Dim r As Long, c As Long

    lastResultRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    Set copyRange = ws.Range("A1:B" & lastResultRow)
    Set pasteRange = ws.Range("C1:D" & lastResultRow)

    copyRange.Copy
    pasteRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' === Apply cell borders to C and D columns ===
    For r = 1 To lastResultRow
        For c = 3 To 4 ' Columns C and D
            With ws.Cells(r, c).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        Next c
    Next r

    ' === Set font properties based on column A ===
    With ws.Range("C1:D" & lastResultRow)
        .Font.Name = ws.Range("A1").Font.Name
        .Font.Size = ws.Range("A1").Font.Size
        .Font.Bold = ws.Range("A1").Font.Bold
    End With

    MsgBox "? Python integration completed successfully.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
