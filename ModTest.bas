Attribute VB_Name = "ModTest"

    
Public Sub TestDealCalc()
    Dim MailInbox As ClsMailInbox
        
    Set MailSystem = New ClsMailSystem
    Set MailInbox = New ClsMailInbox
    With MailInbox
        .MailFolder = "contact@cbs-capital.co.uk"
    End With
    
    MailInbox.SendDealCalc "jules.turner@hotmail.co.uk", "James"
    
    Set MailSystem = Nothing
    Set MailInbox = Nothing

End Sub

Public Sub New_Mail()

Dim oAccount As Account
Dim oMail As MailItem

For Each oAccount In Session.Accounts
    Debug.Print oAccount
    If LCase(oAccount) = LCase("text copied from the immediate window") Then
        Set oMail = CreateItem(olMailItem)
        oMail.SendUsingAccount = oAccount
        oMail.Display
    End If
    Next

ExitRoutine:
    Set oMail = Nothing

End Sub
