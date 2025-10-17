Attribute VB_Name = "Module1"
' =============================================
' Project : Excel VBA Automation
' Module  : Row_Search_Highlight.bas
' Author  : Osman Uluhan
' Date    : 2025-10-14
' Version : 1.0
' =============================================
'
' Description:
' Searches for a keyword typed in a specific cell (like B1)
' and highlights all rows containing that keyword.
'
' Features:
' - Dynamic search for text or numbers
' - Highlights matching rows
' - Clears previous highlights automatically
' - Fully automatic; no button needed (event-based)
'
' License:
' MIT License – Free to use and modify with attribution
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =============================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim SearchCell As Range
    Dim DataRange As Range
    Dim c As Range
    Dim SearchText As String
    
    ' Arama hücresi
    Set SearchCell = Me.Range("B1")
    If Intersect(Target, SearchCell) Is Nothing Then Exit Sub
    
    ' Veri aralýðý (gerekirse geniþlet)
    Set DataRange = Me.Range("A3:E50")
    DataRange.Interior.ColorIndex = xlNone  ' önceki renklendirmeyi temizle
    
    ' Arama metni
    SearchText = Trim(SearchCell.Value)
    If SearchText = "" Then Exit Sub
    
    ' Her hücreyi tek tek kontrol et
    For Each c In DataRange
        ' Tam eþleþme kontrolü (büyük/küçük harf fark etmez)
        If StrComp(c.Text, SearchText, vbTextCompare) = 0 Then
            c.Interior.Color = RGB(255, 255, 153) ' sadece eþleþen hücre sarý
        End If
    Next c
End Sub
