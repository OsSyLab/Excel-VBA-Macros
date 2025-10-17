Attribute VB_Name = "Module1"
' =======================================================
' Project : Excel VBA Automation
' Module  : Quotation_Form_Calculator.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 1.0 (Stable)
' =======================================================
'
' Description:
' Calculates subtotal, tax, and total for a quotation form.
' Ideal for automated invoice or proposal templates.
'
' -------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -------------------------------------------------------
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =======================================================

Sub CalculateQuote()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim taxRate As Double
    Dim subtotal As Double
    Dim totalTax As Double
    Dim grandTotal As Double
    
    Set ws = ThisWorkbook.Sheets("Quotation")
    taxRate = 0.18 ' 18% VAT
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    subtotal = 0
    totalTax = 0
    grandTotal = 0
    
    For i = 2 To lastRow - 1
        If ws.Cells(i, "B").Value <> "" Then
            ws.Cells(i, "E").Value = ws.Cells(i, "C").Value * ws.Cells(i, "D").Value
            ws.Cells(i, "F").Value = ws.Cells(i, "E").Value * taxRate
            ws.Cells(i, "G").Value = ws.Cells(i, "E").Value + ws.Cells(i, "F").Value
            
            subtotal = subtotal + ws.Cells(i, "E").Value
            totalTax = totalTax + ws.Cells(i, "F").Value
            grandTotal = grandTotal + ws.Cells(i, "G").Value
        End If
    Next i
    
    ws.Cells(lastRow, "E").Value = subtotal
    ws.Cells(lastRow, "F").Value = totalTax
    ws.Cells(lastRow, "G").Value = grandTotal
    
    MsgBox "? Quotation calculated successfully!", vbInformation
End Sub

Sub ClearForm()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Quotation")
    
    ' --- Dinamik form alanýný tespit et ---
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then lastRow = 10
    
    ' --- Sadece kullanýcý veri giriþ alanlarýný temizle (B:D) ---
    ws.Range("B2:D" & lastRow).ClearContents
    
    ' --- Hesaplama alanlarýný temizle (E2:G...) ama baþlýklara dokunma ---
    ws.Range("E2:G" & lastRow).ClearContents
    
    ' --- Baþlýklarý yeniden güvence altýna al (isteðe baðlý koruma) ---
    If ws.Cells(1, "E").Value = "" Then ws.Cells(1, "E").Value = 0
    If ws.Cells(1, "F").Value = "" Then ws.Cells(1, "F").Value = 0
    If ws.Cells(1, "G").Value = "" Then ws.Cells(1, "G").Value = 0
    
    MsgBox "?? Form cleared successfully! Headers preserved.", vbInformation
End Sub
