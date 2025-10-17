Attribute VB_Name = "Module1"
' =======================================================
' Project : Excel VBA Automation
' Module  : Export_PDF_Report.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 1.1 (Centered PDF Layout)
' =======================================================
'
' Description:
' Exports the visible (filtered) data from the active worksheet
' as a centered and properly formatted PDF file.
' Automatically creates a folder and names the file based on
' the selected department (B2) and current date.
'
' -------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -------------------------------------------------------
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =======================================================

Sub ExportToPDF()
    Dim ws As Worksheet
    Dim pdfName As String, folderPath As String
    Dim depName As String, todayDate As String
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' --- Page setup for top-centered PDF output ---
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = False   ' Align slightly higher instead of full center
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.3)    ' smaller top margin
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    
    ' --- Get department info (from cell B2) ---
    depName = "All"
    On Error Resume Next
    depName = ws.Range("B2").Value
    On Error GoTo 0
    
    ' --- Prepare folder and file name ---
    todayDate = Format(Date, "yyyy-mm-dd")
    folderPath = Environ("USERPROFILE") & "\Desktop\Reports\"
    
    ' Create folder if it doesn't exist
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    ' Generate dynamic PDF name
    pdfName = folderPath & "Report_" & depName & "_" & todayDate & ".pdf"
    
    ' --- Export to PDF ---
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' --- Confirmation message ---
    MsgBox "? Report exported successfully to:" & vbCrLf & pdfName, _
           vbInformation, "PDF Export Complete"
End Sub
