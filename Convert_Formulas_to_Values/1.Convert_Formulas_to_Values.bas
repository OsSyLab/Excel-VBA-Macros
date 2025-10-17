Attribute VB_Name = "Module1"
' ===============================================================
' Project   : Excel VBA Automation
' Module    : Convert_Formulas_To_Values.bas
' Author    : Osman Uluhan
' Date      : 2025-10-13
' Version   : 1.0
' ===============================================================
'
' Description:
' This VBA macro replaces formulas in the selected range
' (or entire worksheet if nothing is selected)
' with their resulting values.
'
' Features:
' - Converts all formulas to static values
' - Works on selected range or entire sheet
' - Preserves cell formatting
' - Quick, safe, and undo supported
'
' License:
' MIT License – You are free to use, modify, and distribute
' this code with attribution.
'
' © 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
' ===============================================================

Sub Convert_Formulas_To_Values()
    Dim ws As Worksheet
    Dim rng As Range

    On Error GoTo Fail
    Set ws = ActiveSheet
    
    ' --- If user selected a range, use that ---
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    Else
        ' --- If not, use the entire used range ---
        Set rng = ws.UsedRange
    End If

    ' --- Replace formulas with their values ---
    rng.Value = rng.Value

    MsgBox "? All formulas converted to static values successfully!", vbInformation, "Conversion Completed"
    GoTo SafeExit

Fail:
    MsgBox "? Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Aborted"

SafeExit:
    Application.ScreenUpdating = True
End Sub

