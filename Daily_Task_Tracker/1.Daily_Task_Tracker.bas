Attribute VB_Name = "Module1"
'==============================================
' Project : Excel VBA Automation
' Module  : Daily_Task_Tracker.bas
' Author  : Osman Uluhan
' Date    : 2025-10-16
' Version : 1.0
'==============================================
' Description :
' This macro updates task progress and color-codes rows
' based on task status (Pending / In Progress / Completed).
' It also calculates the average completion rate.
'==============================================

Sub UpdateProgress()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalProgress As Double
    Dim countTasks As Long
    Dim avgProgress As Double
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("TaskTracker")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    totalProgress = 0
    countTasks = 0
    
    For i = 2 To lastRow
        Set rng = ws.Range("A" & i & ":E" & i) ' ?? ' Only between A:E cells are painting
        
        Select Case LCase(Trim(ws.Cells(i, "C").Value))
            Case "completed"
                ws.Cells(i, "D").Value = 1
                rng.Interior.Color = RGB(198, 239, 206) ' Light green
            Case "in progress"
                ws.Cells(i, "D").Value = 0.5
                rng.Interior.Color = RGB(255, 235, 156) ' Light yellow
            Case "pending"
                ws.Cells(i, "D").Value = 0
                rng.Interior.Color = RGB(242, 242, 242) ' Light gray
            Case Else
                rng.Interior.Color = xlNone
        End Select
        
        If ws.Cells(i, "C").Value <> "" Then
            totalProgress = totalProgress + ws.Cells(i, "D").Value
            countTasks = countTasks + 1
        End If
    Next i
    
    If countTasks > 0 Then
        avgProgress = totalProgress / countTasks
        ws.Range("F2").Value = Format(avgProgress * 100, "0") & "%"
    Else
        ws.Range("F2").Value = "0%"
    End If
    
    MsgBox "? Task progress updated successfully!", vbInformation
End Sub

'==============================================
' Clears task list but keeps headers
'==============================================
Sub ClearTasks()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaskTracker")
    
    ws.Range("A2:E100").ClearContents
    ws.Range("A2:E100").Interior.Color = xlNone
    ws.Range("F2").Value = ""
    
    MsgBox "?? Task list cleared successfully!", vbInformation
End Sub
