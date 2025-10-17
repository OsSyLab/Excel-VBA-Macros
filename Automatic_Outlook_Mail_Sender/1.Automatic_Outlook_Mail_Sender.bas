Attribute VB_Name = "Module1"
' =======================================================
' Project : Excel VBA Automation
' Module  : SendMail.bas
' Author  : Osman Uluhan
' Date    : 2025-10-15
' Version : 1.1 (Stable - Fixed Sender Account)
' =======================================================
'
' Description:
' Sends automated e-mails through Microsoft Outlook
' using a specific sender account (osmanuluhan@hotmail.com)
' based on data stored in an Excel table.
'
' -------------------------------------------------------
' License:
' MIT License – Free to use, modify, and distribute
' with attribution.
' -------------------------------------------------------
'
' © 2025 Data Solutions Lab. by Osman Uluhan
' =======================================================


Sub SendMail()
    Dim OutlookApp As Object, OutlookMail As Object
    Dim OutlookAccount As Object, ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim mailTo As String, mailSubject As String, mailBody As String
    Dim senderFound As Boolean

    ' Connect to Outlook
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Find your Outlook account
    For Each OutlookAccount In OutlookApp.Session.Accounts
        If LCase(OutlookAccount.SmtpAddress) = "your_email@example.com" Then
            senderFound = True
            Exit For
        End If
    Next OutlookAccount

    ' Reference sheet
    Set ws = ThisWorkbook.Sheets("MailList")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Send loop
    For i = 2 To lastRow
        mailTo = ws.Cells(i, "B").Value
        mailSubject = ws.Cells(i, "C").Value
        mailBody = ws.Cells(i, "D").Value
        If mailTo <> "" Then
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .SendUsingAccount = OutlookAccount
                .To = mailTo
                .Subject = mailSubject
                .Body = mailBody
                .Send
            End With
            ws.Cells(i, "E").Value = "Sent"
        Else
            ws.Cells(i, "E").Value = "Missing Email"
        End If
    Next i
End Sub
