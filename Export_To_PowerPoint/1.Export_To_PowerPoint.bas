Attribute VB_Name = "Module1"
'============================================================
' Project : Excel VBA Automation
' Module  : Export_PowerPoint_Report.bas
' Author  : Osman Uluhan
' Date    : 2025-10-17
' Version : 1.0
'============================================================
' Description:
' This macro exports a selected range or the used range
' from the active worksheet into a new PowerPoint presentation.
' It automatically creates a new slide, pastes the table as an image,
' centers it neatly, and saves the presentation on the Desktop.
'============================================================

Sub ExportToPowerPoint()

    Dim ppApp As Object
    Dim ppPres As Object
    Dim ppSlide As Object
    Dim ws As Worksheet
    Dim exportRange As Range
    Dim filePath As String
    
    On Error Resume Next
    Set ppApp = GetObject(Class:="PowerPoint.Application")
    If ppApp Is Nothing Then
        Set ppApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    If ppApp Is Nothing Then
        MsgBox "?? PowerPoint is not installed or accessible.", vbCritical
        Exit Sub
    End If
    
    ppApp.Visible = True
    
    '--- Get data ---
    Set ws = ActiveSheet
    On Error Resume Next
    Set exportRange = Application.InputBox( _
        Prompt:="Select the range to export to PowerPoint:", _
        Title:="Export Range", Type:=8)
    On Error GoTo 0
    
    If exportRange Is Nothing Then
        MsgBox "? Export cancelled.", vbExclamation
        Exit Sub
    End If
    
    '--- Copy the range as a picture ---
    exportRange.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    '--- Create PowerPoint file ---
    Set ppPres = ppApp.Presentations.Add
    Set ppSlide = ppPres.Slides.Add(1, 1) ' ppLayoutTitle
    
    '--- Paste the image ---
    ppSlide.Shapes.Paste.Select
    With ppApp.ActiveWindow.Selection.ShapeRange
        ' Center horizontally & vertically
        .Left = (ppApp.ActivePresentation.PageSetup.SlideWidth - .Width) / 2
        .Top = (ppApp.ActivePresentation.PageSetup.SlideHeight - .Height) / 2
    End With
    
    '--- Add a slide title ---
    ppSlide.Shapes.Title.TextFrame.TextRange.Text = "Excel PowerPoint Report"
    ppSlide.Shapes.Title.TextFrame.TextRange.Font.Size = 32
    ppSlide.Shapes.Title.TextFrame.TextRange.Font.Bold = True
    
    '--- Save file to Desktop ---
    filePath = Environ("USERPROFILE") & "\Desktop\Report_AutoExport.pptx"
    ppPres.SaveAs filePath
    
    MsgBox "? PowerPoint report generated successfully!" & vbCrLf & _
           "File saved to: " & filePath, vbInformation, "Export Complete"

End Sub

