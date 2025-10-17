Attribute VB_Name = "Module1"
' =======================================================
' Project : Excel VBA Automation
' Module  : Export_Dashboard_Refresh.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 1.0 (Stable)
' =======================================================
'
' Description:
' Automatically refreshes all pivot tables, query tables,
' and Power Query connections in the workbook.
' Designed for dashboards that require one-click updates.
'
' Features:
' - Refreshes all pivot tables across all sheets
' - Updates all data connections and Power Query links
' - Rebuilds all calculations (full recalc)
' - Displays a confirmation message when complete
'
' Recommended Usage:
' Attach to a “Refresh Dashboard” button or
' call automatically on Workbook_Open.
'
' -------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -------------------------------------------------------
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =======================================================


Sub RefreshDashboard()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim qt As QueryTable
    Dim co As WorkbookConnection
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing dashboard..."
    
    ' === Refresh all pivot tables ===
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws
    
    ' === Refresh all query tables ===
    For Each ws In ThisWorkbook.Worksheets
        For Each qt In ws.QueryTables
            qt.Refresh BackgroundQuery:=False
        Next qt
    Next ws
    
    ' === Refresh all Power Query connections ===
    For Each co In ThisWorkbook.Connections
        On Error Resume Next
        co.Refresh
        On Error GoTo 0
    Next co
    
    ' === Recalculate everything ===
    Application.CalculateFullRebuild
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "? Dashboard updated successfully!", vbInformation, "Dashboard Refresh"
End Sub
