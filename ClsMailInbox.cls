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
Private ObjectNS As Outlook.Namespace
Private FldrInbox As Outlook.Folder

' ===============================================================
' OutlookApp_NewMailEx
' Event on new email
' ---------------------------------------------------------------
Private Sub OutlookApp_NewMailEx(ByVal EntryIDCollection As String)
    Dim Msg As Outlook.MailItem
    Dim StoreID As String
    
    On Error GoTo ErrorHandler
        
    StoreID = FldrInbox.StoreID
    Set Msg = ObjectNS.GetItemFromID(EntryIDCollection, StoreID)
    
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
    Set FldrInbox = ObjectNS.Folders("contact@cbs-capital.co.uk").Folders("Inbox")
    Set InboxItems = FldrInbox.Items
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
                
                If DB Is Nothing Then DBConnect
                
                Set RstContact = ModDatabase.SQLQuery("SELECT * FROM TblContact WHERE EmailAddress = '" & Email & "'")
                Set RstContMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) As MaxNo FROM TblContact")
                
                If RstContact.RecordCount = 0 Then
                    Dim NextNo As Integer
                    NextNo = RstContMaxNo!MaxNo + 1
                    
                    DB.Execute "INSERT INTO TblContact (ContactNo, ContactType, ContactName, EmailAddress) " _
                    & "VALUES (" & NextNo & ", 'Lead'" & ", '" & FirstName & " " & LastName & "', '" & Email & " ') "
                End If
            End If
        End If
    End With
    Set RstContact = Nothing
    Set RstContMaxNo = Nothing
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set OutlookApp = Outlook.Application
    Set ObjectNS = OutlookApp.GetNamespace("MAPI")
    Set FldrInbox = ObjectNS.Folders("contact@cbs-capital.co.uk").Folders("Inbox")
    Set InboxItems = FldrInbox.Items
    ParseEmails
End Sub
' ---------------------------------------------------------------
Private Sub Class_Terminate()
    Set OutlookApp = Nothing
    Set ObjectNS = Nothing
    Set FldrInbox = Nothing
    Set InboxItems = Nothing

End Sub
' ===============================================================

