Attribute VB_Name = "Module1"
'===============================================================
' Project : Excel VBA Automation
' Module  : Data_Summary.bas
' Author  : Osman Uluhan
' Date    : 2025-10-17
' Version : 1.0
'
' Description:
'   This macro automatically groups data by category (first column)
'   and sums numerical values (second column), creating a new
'   summary sheet with clean, formatted output.
'
'===============================================================

Option Explicit

Sub GenerateSummary()
    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim dict As Object
    Dim key As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim grpName As String
    Dim grpValue As Double
    Dim summaryRow As Long
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ' Delete old summary if exists
    Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new summary sheet
    Set wsSummary = ThisWorkbook.Worksheets.Add
    wsSummary.Name = "Summary"
    
    ' Set data source sheet
    Set wsData = ThisWorkbook.Worksheets("DataSummary")
    
    ' Find last data row
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Use Dictionary to group data
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        grpName = Trim(wsData.Cells(i, 1).Value)
        If grpName <> "" Then
            grpValue = wsData.Cells(i, 2).Value
            If dict.Exists(grpName) Then
                dict(grpName) = dict(grpName) + grpValue
            Else
                dict.Add grpName, grpValue
            End If
        End If
    Next i
    
    ' Write header
    wsSummary.Range("A1").Value = "Department"
    wsSummary.Range("B1").Value = "Total Amount"
    wsSummary.Range("A1:B1").Font.Bold = True
    wsSummary.Range("A1:B1").Interior.Color = RGB(189, 215, 238)
    
    ' Write summarized data
    summaryRow = 2
    For Each key In dict.Keys
        wsSummary.Cells(summaryRow, 1).Value = key
        wsSummary.Cells(summaryRow, 2).Value = dict(key)
        summaryRow = summaryRow + 1
    Next key
    
    ' Auto format
    With wsSummary
        .Columns("A:B").AutoFit
        .Range("A1:B" & summaryRow - 1).Borders.LineStyle = xlContinuous
        .Range("B2:B" & summaryRow - 1).NumberFormat = "#,##0"
    End With
    
    MsgBox "? Data summarized successfully!" & vbCrLf & _
           "Summary sheet created: 'Summary'", vbInformation, "Summary Complete"

End Sub

