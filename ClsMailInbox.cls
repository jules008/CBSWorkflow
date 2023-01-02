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
Private WithEvents InboxItems As Outlook.Items
Attribute InboxItems.VB_VarHelpID = -1
Private WithEvents OutlookApp As Outlook.Application
Attribute OutlookApp.VB_VarHelpID = -1
Private ObjectNS As Outlook.Namespace
Private pFldrInbox As Outlook.Folder
Private pActive As Boolean
Private pMailFolder As String

' ===============================================================
Public Property Get MailFolder() As String
    MailFolder = pMailFolder
End Property
' ---------------------------------------------------------------
Public Property Let MailFolder(ByVal vNewValue As String)
    On Error Resume Next
    
    pMailFolder = vNewValue
    Set pFldrInbox = ObjectNS.Folders(pMailFolder).Folders("Inbox")
    
    If pFldrInbox Is Nothing Then
        pActive = False
    Else
        pActive = True
        Set InboxItems = pFldrInbox.Items
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
Private Sub OutlookApp_NewMailEx(ByVal EntryIDCollection As String)
    Dim Msg As Outlook.MailItem
    Dim StoreID As String
    
    On Error GoTo ErrorHandler
        
    StoreID = pFldrInbox.StoreID
    Set Msg = ObjectNS.GetItemFromID(EntryIDCollection, StoreID)
    
    Debug.Print "Event triggered"
    
    CheckMail Msg
    
ExitNewItem:

    Exit Sub

ErrorHandler:
    
    Resume ExitNewItem
End Sub

' ===============================================================
' OutlookApp_Quit
' Necessary for microsoft bug that prevents startup running next time
' ---------------------------------------------------------------
Private Sub OutlookApp_Quit()
End Sub

' ===============================================================
' OutlookApp_Startup
' Loads necessary Objects
' ---------------------------------------------------------------
Private Sub OutlookApp_Startup()
    Set OutlookApp = Outlook.Application
    Set ObjectNS = OutlookApp.GetNamespace("MAPI")
    Set pFldrInbox = ObjectNS.Folders("contact@cbs-capital.co.uk").Folders("Inbox")
    Set InboxItems = pFldrInbox.Items
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
        Debug.Print i
        i = i + 1
    Next
End Sub

' ===============================================================
' CheckMail
' Checks and imports emails
' ---------------------------------------------------------------
Public Sub CheckMail(Msg As Outlook.MailItem)
    Dim FirstName As String
    Dim EmailBody() As String
    Dim LastName As String
    Dim Email As String
    Dim RstContact As Recordset
    Dim RstContMaxNo As Recordset
    Dim i As Integer
    
    With Msg
        If Msg.UnRead = True Then
            If InStr(1, .Subject, "CBS DB Import", vbTextCompare) Then
                
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
                
                Debug.Print "Parsed email", FirstName, LastName, Email
                
                SendDealCalc Email, FirstName
                
                If DB Is Nothing Then DBConnect
                
                Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE EmailAddress = '" & Email & "'")
                Set RstContMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) As MaxNo FROM TblContact")
                
                If RstContact.RecordCount = 0 Then
                    Debug.Print "Not in database"
                    Dim NextNo As Integer
                    
                    If Not IsNull(RstContMaxNo!MaxNo) Then
                    NextNo = RstContMaxNo!MaxNo + 1
                    Else
                        NextNo = 1
                    End If
                    
                    DB.Execute "INSERT INTO TblContact (ContactNo, ContactType, ContactName, EmailAddress, ComFrq, LastComm) " _
                    & "VALUES (" & NextNo & ", 'Lead'" & ", '" & FirstName & " " & LastName & "', '" & Email & " ', " & 30 & " , '" & Format(Now, "dd mm yy") & " ') "
                    
                End If
            End If
        End If
    End With
    Set RstContact = Nothing
    Set RstContMaxNo = Nothing
End Sub

' ===============================================================
' SendDealCalc
' sends email with deal calc attached
' ---------------------------------------------------------------
Public Sub SendDealCalc(EmailAdd As String, Name As String)
    Dim DealEmail As ClsMailItem
    Dim MailAccount As Outlook.Account
    
    Set MailSystem = Nothing
    Set MailSystem = New ClsMailSystem
    Set MailAccount = Session.Accounts("contact@cbs-capital.co.uk")
    
    With MailSystem.MailItem
        .To = EmailAdd
        .Subject = "Deal Calculator"
        .Attachments.Add GetDocLocalPath(ThisWorkbook.Path) & TEMPLATE_STORE & DEAL_CALC_FILE & ".xlsx"
        .Body = Name & vbCr _
            & "Thanks for taking the time to download your very own CBS Deal Calculator. " & vbCr & vbCr _
            & "This calculator will enable you to quickly assess projects to ensure they are hitting the right profit levels. " & vbCr & vbCr _
            & "Your project should ideally be hitting 28%+ profit on GDV (pre finance costs). " & vbCr & vbCr _
            & "If your project is scoring low, don’t worry, still book a call to see how our structured finance solutions can help you." & vbCr & vbCr _
            & "Lets Talk" & vbCr & vbCr _
            & "https://calendly.com/cbscapital/heather" & vbCr & vbCr _
            & "Heather Critchlow" & vbCr _
            & "Partner" & vbCr _
            & "CBS Capital" & vbCr _
            & "W| www.cbscapital.co.uk"
            .SendUsingAccount = MailAccount
    End With
    With MailSystem
        .SendEmail
        Debug.Print "email Sent"
    End With
    
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set OutlookApp = Outlook.Application
    Set ObjectNS = OutlookApp.GetNamespace("MAPI")
End Sub
' ---------------------------------------------------------------
Private Sub Class_Terminate()
    Set OutlookApp = Nothing
    Set ObjectNS = Nothing
    Set pFldrInbox = Nothing
    Set InboxItems = Nothing

End Sub
' ===============================================================

