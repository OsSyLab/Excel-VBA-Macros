Attribute VB_Name = "Module1"
Sub Conditional_Formatting()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    
    ' Select the active sheet
    Set ws = ActiveSheet
    
    ' Define the range to format (for example column C)
    Set dataRange = ws.Range("C2:C100")
    
    ' Clear any existing formatting
    dataRange.Interior.ColorIndex = xlNone
    
    ' Loop through each cell and apply color based on value
    For Each cell In dataRange
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            If cell.Value > 0 Then
                cell.Interior.Color = RGB(198, 239, 206) ' light green
            ElseIf cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 199, 206) ' light red
            Else
                cell.Interior.Color = RGB(255, 235, 156) ' yellow
            End If
        End If
    Next cell

    MsgBox "? Conditional formatting applied successfully!", vbInformation
End Sub

