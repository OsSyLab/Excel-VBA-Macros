Attribute VB_Name = "Module1"
' =============================================
' Project : Excel VBA Automation
' Module  : Protect_Unprotect_Sheet.bas
' Author  : Osman Uluhan
' Date    : 2025-10-14
' Version : 1.0
' =============================================
'
' Description:
' This VBA macro protects the entire sheet while allowing
' specific cells to remain editable. It can also unprotect
' the sheet easily when needed.
'
' Features:
' - Locks all cells in the active worksheet
' - Keeps only specified range unlocked (editable)
' - Option to protect or unprotect instantly
' - Prevents accidental formula or data changes
'
' License:
' MIT License – You are free to use, modify, and distribute
' this code with attribution.
'
' © 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
' =============================================

Sub Protect_SelectedSheet()
    Dim ws As Worksheet
    Dim editableRange As Range
    Dim pwd As String
    
    ' === Define worksheet ===
    Set ws = ActiveSheet
    
    ' === Set your editable area (change as needed) ===
    Set editableRange = ws.Range("B2:D10")
    pwd = "secure123"   ' <-- You can change the password here
    
    ' === Unlock editable range first ===
    ws.Cells.Locked = True
    editableRange.Locked = False
    
    ' === Protect the sheet ===
    ws.Protect Password:=pwd, UserInterfaceOnly:=True
    MsgBox "? Sheet is now protected, only range " & editableRange.Address & " is editable.", vbInformation
End Sub


Sub Unprotect_SelectedSheet()
    Dim ws As Worksheet
    Dim pwd As String
    
    Set ws = ActiveSheet
    pwd = "secure123"   ' Must match the password above
    
    ' === Unprotect the sheet ===
    ws.Unprotect Password:=pwd
    MsgBox "?? Sheet is now unprotected.", vbInformation
End Sub

