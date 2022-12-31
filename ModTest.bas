Attribute VB_Name = "ModTest"

    
Public Sub TestDealCalc()
    Dim MailInbox As ClsMailInbox
        
    Set MailSystem = New ClsMailSystem
    Set MailInbox = New ClsMailInbox
    With MailInbox
        .MailFolder = "contact@cbs-capital.co.uk"
    End With
    
    MailInbox.SendDealCalc "jules.turner@hotmail.co.uk"
    
    Set MailSystem = Nothing
    Set MailInbox = Nothing

End Sub
