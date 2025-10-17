Attribute VB_Name = "Module1"
' ============================================================
'  Project : Excel VBA Automation
'  Module  : One_Click_Export_Pro.bas
'  Author  : Osman Uluhan
'  Date    : 2025-10-12
'  Version : 1.1
' ------------------------------------------------------------
'  Description:
'  This VBA macro exports the active sheet as a professional-
'  looking PDF report with formatted header, footer, and date.
'  The output is automatically saved to the user's Desktop.
'
'  Features:
'   - Adds header with sheet name and current date
'   - Adds footer with page numbers and copyright
'   - Fits sheet to one page width (no layout distortion)
'   - Automatically saves PDF to Desktop
'   - Optional: can include company logo (future update)
'
'  License:
'  MIT License – You are free to use, modify, and distribute
'  this code with attribution.
'
'  © 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.
' ============================================================

Sub One_Click_Export_Pro()
    Dim ws As Worksheet
    Dim filePath As String
    Dim exportName As String
    Dim currentDate As String
    
    ' --- Select the active sheet
    Set ws = ActiveSheet
    
    ' --- Get current date
    currentDate = Format(Date, "dd mmmm yyyy")
    
    ' --- Create file name (SheetName_Report_Date)
    exportName = ws.Name & "_Report_" & Format(Now, "yyyy-mm-dd_hhmm")
    
    ' --- Define save path (Desktop)
    filePath = Environ("USERPROFILE") & "\Desktop\" & exportName & ".pdf"
    
    ' --- Page setup configuration
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .CenterVertically = False
        
        ' Margins
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.7)
        .BottomMargin = Application.InchesToPoints(0.7)
        
        ' --- Header (top section)
        .CenterHeader = "&""Calibri,Bold""&14 " & ws.Name & " Report"
        .LeftHeader = "&""Calibri""&10 Generated: " & currentDate
        
        ' --- Footer (bottom section)
        .CenterFooter = "&""Calibri""&10 Page &P of &N"
        .RightFooter = "&""Calibri""&10 © Data Solutions Lab. by Osman Uluhan"
    End With
    
    ' --- Export as PDF
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath
    On Error GoTo 0
    
    ' --- Confirmation message
    MsgBox "? PDF exported successfully!" & vbCrLf & _
           "Saved on Desktop as:" & vbCrLf & filePath, _
           vbInformation, "Export Completed"
End Sub


