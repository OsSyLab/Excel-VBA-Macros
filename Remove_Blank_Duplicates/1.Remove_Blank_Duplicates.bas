Attribute VB_Name = "Module1"
' ============================================================
'  Project : Excel VBA Automation
'  Module  : Remove_Blank_Duplicates.bas
'  Author  : Osman Uluhan
'  Date    : 2025-10-13
'  Version : 1.0
' ------------------------------------------------------------
'  Description:
'  This VBA macro cleans up data by removing blank rows
'  and duplicate entries from a selected range or worksheet.
'
'  Features:
'   - Deletes fully blank rows
'   - Removes duplicate records based on entire row values
'   - Works on active sheet or selected range
'   - Keeps the first occurrence of duplicates
'
'  License:
'  MIT License – You are free to use, modify, and distribute
'  this code with attribution.
'
'  © 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
' ============================================================

Sub Clean_Blanks_And_RemoveDuplicates()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lo As ListObject
    Dim r As Long
    Dim arr As Variant
    
    On Error GoTo Fail
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' === 1?? Tablon varsa ===
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        
        ' --- Boþ satýrlarý sil ---
        If Not lo.DataBodyRange Is Nothing Then
            For r = lo.DataBodyRange.Rows.Count To 1 Step -1
                If WorksheetFunction.CountA(lo.DataBodyRange.Rows(r)) = 0 Then
                    lo.ListRows(r).Delete
                End If
            Next r
        End If
        
        ' --- Dedup (sabit sütun sayýsý kullanýyoruz; tablo 4 sütunlu) ---
        lo.Range.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
        
    Else
        ' === 2?? Tablon yoksa ===
        Set rng = ws.Range("A1").CurrentRegion
        
        ' --- Boþ satýrlarý sil ---
        For r = rng.Rows.Count To 2 Step -1
            If WorksheetFunction.CountA(rng.Rows(r)) = 0 Then
                rng.Rows(r).EntireRow.Delete
            End If
        Next r
        
        ' --- Dedup ---
        Set rng = ws.Range("A1").CurrentRegion
        rng.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
    End If
    
    MsgBox "? Boþ satýrlar silindi ve yinelenenler kaldýrýldý!", vbInformation
    GoTo SafeExit

Fail:
    MsgBox "? Hata " & Err.Number & ": " & Err.Description, vbCritical, "Macro aborted"

SafeExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
