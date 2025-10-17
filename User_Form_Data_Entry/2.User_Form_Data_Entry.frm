VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Data Entry"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "User_Form_Data_Entry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
    Dim ws As Worksheet
    Dim nextRow As Long
    
    ' Kayýt yazýlacak sayfa
    Set ws = ThisWorkbook.Worksheets("Database")
    
    ' Ýlk boþ satýr
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If nextRow = 1 And ws.Cells(1, 1).Value <> "" Then
        nextRow = 2
    Else
        nextRow = nextRow + 1
    End If
    
    ' Basit doðrulama
    If Trim(Me.txtName.Value) = "" Or Trim(Me.txtEmail.Value) = "" Then
        MsgBox "Name ve Email boþ olamaz.", vbExclamation
        Exit Sub
    End If
    
    ' Veriyi yaz
    ws.Cells(nextRow, 1).Value = Me.txtName.Value
    ws.Cells(nextRow, 2).Value = Me.txtEmail.Value
    ws.Cells(nextRow, 3).Value = Me.txtPhone.Value
    ws.Cells(nextRow, 4).Value = Me.txtCity.Value
    ws.Cells(nextRow, 5).Value = Now
    
    MsgBox "1 kayýt eklendi ?", vbInformation
    
    ' Temizle
    Me.txtName.Value = ""
    Me.txtEmail.Value = ""
    Me.txtPhone.Value = ""
    Me.txtCity.Value = ""
    Me.txtName.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub txtName_Change()

End Sub

Private Sub UserForm_Click()

End Sub
