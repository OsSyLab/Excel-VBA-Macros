Attribute VB_Name = "Module1"
' ============================================
' Project : Excel VBA Automation
' Module  : Auto_Insert_Date.bas
' Author  : Osman Uluhan
' Date    : 2025-10-13
' Version : 1.0
' ============================================

' Description:
' Inserts today’s date automatically in column D
' when any cell in column C is filled or changed.
' Works on all Excel versions, no manual run required.

Option Explicit

Sub Auto_Insert_Date()
    Dim ws As Worksheet
    Dim DateCol As Long
    Dim r As Range

    Set ws = ActiveSheet
    DateCol = 4 ' D sütunu

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each r In ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        If ws.Cells(r.Row, "A").Value <> "" And ws.Cells(r.Row, DateCol).Value = "" Then
            ws.Cells(r.Row, DateCol).Value = Date
            ws.Cells(r.Row, DateCol).NumberFormat = "dd.mm.yyyy"
        End If
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "? Dates inserted successfully!", vbInformation, "Auto Date Complete"
End Sub

