VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsMailInbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsMailInbox
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Nov 22
'===============================================================
Option Explicit
Private WithEvents InboxItems As outlook.Items
Attribute InboxItems.VB_VarHelpID = -1
Private WithEvents pOutlookApp As outlook.Application
Attribute pOutlookApp.VB_VarHelpID = -1
Private pObjectNS As outlook.Namespace
Private pFldrInbox As outlook.Folder
Private pActive As Boolean
Private pMailFolder As String
Private WithEvents pMailItem As outlook.MailItem
Attribute pMailItem.VB_VarHelpID = -1

' ===============================================================
Public Property Get MailFolder() As String
    MailFolder = pMailFolder
End Property
' ---------------------------------------------------------------
Public Property Let MailFolder(ByVal vNewValue As String)
    On Error Resume Next
    
    pMailFolder = vNewValue
    Set pFldrInbox = pObjectNS.Folders(pMailFolder).Folders("Inbox")
    
    If pFldrInbox Is Nothing Then
        pActive = False
    Else
        pActive = True
        Set InboxItems = pFldrInbox.Items
        
        If Not LogEvent("Parsing emails", False) Then Err.Raise HANDLED_ERROR

        ParseEmails
    End If
End Property
' ===============================================================
Public Property Get Active() As Boolean
    Active = pActive
End Property

' ===============================================================
' OutlookApp_NewMailEx
' Event on new email
' ---------------------------------------------------------------
Private Sub pOutlookApp_NewMailEx(ByVal EntryIDCollection As String)
    Dim Msg As outlook.MailItem
    Dim StoreID As String
    
    On Error GoTo ExitNewItem
        
    StoreID = pFldrInbox.StoreID
    
    Set Msg = pObjectNS.GetItemFromID(EntryIDCollection, StoreID)
    
    If Not LogEvent("New email arrived", False) Then Err.Raise HANDLED_ERROR
    CheckMail Msg
    
ExitNewItem:

    Exit Sub

ErrorHandler:
    
    Resume ExitNewItem
End Sub

' ===============================================================
' pOutlookApp_Quit
' Necessary for microsoft bug that prevents startup running next time
' ---------------------------------------------------------------
Private Sub pOutlookApp_Quit()
    If Not LogEvent("Outlook Closing", False) Then Err.Raise HANDLED_ERROR
    MsgBox "OUTLOOK IS CLOSING DOWN" & vbCr _
                        & " Emails from potential Clients will no longer be monitored and replied to!!", vbCritical + vbOKOnly, APP_NAME
End Sub

' ===============================================================
' pOutlookApp_Startup
' Loads necessary Objects
' ---------------------------------------------------------------
Private Sub pOutlookApp_Startup()
'    Set pObjectNS = pOutlookApp.GetNamespace("MAPI")
'    Set pFldrInbox = pObjectNS.Folders("contact@cbs-capital.co.uk").Folders("Inbox")
'    Set InboxItems = pFldrInbox.Items
'    If Not LogEvent("Outlook Opened", False) Then Err.Raise HANDLED_ERROR
    ParseEmails
End Sub

' ===============================================================
' ParseEmails
' scan all emails for CBS Imports
' ---------------------------------------------------------------
Public Sub ParseEmails()
    Dim Msg As Object
    Dim i As Integer
    
    For Each Msg In InboxItems
        If TypeName(Msg) = "MailItem" Then
            CheckMail Msg
        End If
        i = i + 1
    Next
End Sub

' ===============================================================
' CheckMail
' Checks and imports emails
' ---------------------------------------------------------------
Public Sub CheckMail(Msg As outlook.MailItem)
    Dim FirstName As String
    Dim EmailBody() As String
    Dim LastName As String
    Dim Email As String
    Dim EnableSend As Boolean
    Dim RstContact As Recordset
    Dim RstContMaxNo As Recordset
    Dim RstEnv As Recordset
    Dim i As Integer
    
    With Msg
        If Msg.UnRead = True Then
            Msg.UnRead = False
            
            If InStr(1, .Subject, "Send Status Report", vbTextCompare) Then
                ModInbox.SendStatusReport
            End If
            
            If InStr(1, .Subject, "Stop Monitor", vbTextCompare) Then
                Stop
            End If
            
            If InStr(1, .Subject, "Start Monitor", vbTextCompare) Then
                ThisWorkbook.Initialise
            End If
            
            
            If InStr(1, .Subject, "CBS DB Import", vbTextCompare) Then
                Msg.UnRead = False
                EmailBody = Split(.Body, vbCr)
                For i = LBound(EmailBody) To UBound(EmailBody)
                    If InStr(1, EmailBody(i), "First Name") > 0 Then
                        FirstName = Trim(Split(EmailBody(i), ":")(1))
                    End If
                    If InStr(1, EmailBody(i), "Last Name") > 0 Then
                        LastName = Trim(Split(EmailBody(i), ":")(1))
                    End If
                    If InStr(1, EmailBody(i), "Email Address") > 0 Then
                        Email = Trim(Split(EmailBody(i), ":")(1))
                    End If
                Next
                
                If Not LogEvent("Parsed email " & FirstName & ", " & LastName & ", " & Email, False) Then Err.Raise HANDLED_ERROR
                
                If DB Is Nothing Then DBConnect
                
                Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE EmailAddress = '" & Email & "'")
                Set RstContMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) As MaxNo FROM TblContact")
                Set RstEnv = ModDatabase.SQLQuery("SELECT Env FROM TblDBVersion")
                
                If RstContact.RecordCount > 0 Then
                    If Not LogEvent("Already in database, so no action", False) Then Err.Raise HANDLED_ERROR
                Else
                    Dim NextNo As Integer
                    
                    If Not IsNull(RstContMaxNo!MaxNo) Then
                    NextNo = RstContMaxNo!MaxNo + 1
                    Else
                        NextNo = 1
                    End If
                    
                    DB.Execute "INSERT INTO TblContact (ContactNo, ContactType, ContactName, EmailAddress, ComFrq, LastComm) " _
                    & "VALUES (" & NextNo & ", 'Lead'" & ", '" & FirstName & " " & LastName & "', '" & Email & " ', " & 30 & " , '" & Format(Now, "dd mm yy") & " ') "
                    
                    If Not LogEvent("Entered into database", False) Then Err.Raise HANDLED_ERROR
                    
                    If RstEnv!Env = "Live" Then EnableSend = True
                    SendDealCalc Email, FirstName, EnableSend
                End If
            End If
        End If
    End With
    Set RstContact = Nothing
    Set RstContMaxNo = Nothing
    Set DB = Nothing
    Set RstEnv = Nothing
End Sub

' ===============================================================
' SendDealCalc
' sends email with deal calc attached
' ---------------------------------------------------------------
Public Sub SendDealCalc(EmailAdd As String, Name As String, EnableSend As Boolean)
    Dim MailItem As MailItem
    Dim MailAccount As outlook.Account
    Dim OLPA As outlook.PropertyAccessor
    Dim AttLogo As outlook.Attachment
    Dim HTML1 As String
    Dim ImgPath As String
    
    ImgPath = GetDocLocalPath(ThisWorkbook.Path) & PICTURES_PATH & LOGO_FILE
    
    Set MailItem = pOutlookApp.CreateItem(0)
    Set MailAccount = Session.Accounts("contact@cbs-capital.co.uk")
    Set AttLogo = MailItem.Attachments.Add(ImgPath)
    Set OLPA = AttLogo.PropertyAccessor
    
    OLPA.SetProperty PR_ATTACH_CONTENT_ID, "Logo"
         
     HTML1 = "<!DOCTYPE html> " _
     & "<html> " _
     & "<BODY style=""line-height: 80%;font-family: Helvetica; font-size: 13px"" > " _
     & "    <p>" & Name & "</p> " _
     & "    <p>Thanks for taking the time to download your very own CBS Deal Calculator.</p>     " _
     & "    <p>This calculator will enable you to quickly assess projects to ensure they are hitting the right profit levels.</p> " _
     & "    <p>Your project should ideally be hitting 28%+ profit on GDV (pre finance costs).</p> " _
     & "    <p>If your project is scoring low, don&#39t worry, still book a call to see how our structured finance solutions can help you.</p>  " _
     & "    <p>Lets Talk</p> " _
     & "    <a href=""https://calendly.com/cbscapital/heather"">https://calendly.com/cbscapital/heather</a> " _
     & "    <p style=""line-height: 100%;"">     " _
     & "    Heather Critchlow<br> " _
     & "    Partner<br> " _
     & "    CBS Capital</p> " _
     & "    W| <a href=""www.cbscapital.co.uk"">www.cbscapital.co.uk</a> " _
     & "    <p><IMG src=""cid:Logo"" width=200></p> " _
     & "    <p style=""font-size: 11px"">Capital is a trading name of&nbsp;the CBS Capital Group of&nbsp;Companies</p> " _
     & "    <p style=""font-size: 11px"">14 | Enterprise House |&nbsp;Cherry Orchard Lane | Salisbury&nbsp;| Wiltshire | SP2 7LD |&nbsp;Registered 11119716,&nbsp;14101720&nbsp;and 13290200</p> " _
     & "    </BODY> " _
     & "</html> "
    With MailItem
        .To = EmailAdd
        .Subject = "Deal Calculator"
        .Attachments.Add GetDocLocalPath(ThisWorkbook.Path) & TEMPLATE_STORE & DEAL_CALC_FILE & ".xlsx"
        .BodyFormat = olFormatHTML
        .HTMLBody = HTML1
        .SendUsingAccount = MailAccount
        
        If EnableSend Then .Send Else .Display
    End With
    
   Set MailItem = Nothing
   Set MailAccount = Nothing
   Set AttLogo = Nothing
   Set OLPA = Nothing
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pOutlookApp = CreateObject(Class:="Outlook.Application")
    Set pObjectNS = pOutlookApp.GetNamespace("MAPI")
End Sub
' ---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pFldrInbox = Nothing
    Set InboxItems = Nothing
    Set pObjectNS = Nothing
    Set pOutlookApp = Nothing

End Sub
' ===============================================================

