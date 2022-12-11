VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsMailSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsMailSystem
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 06 Jul 20
'===============================================================
Option Explicit
Private WithEvents pOutlookApp As Outlook.Application
Attribute pOutlookApp.VB_VarHelpID = -1
Private pReturnMail As ClsMailItem
Private WithEvents pMailItem As Outlook.MailItem
Attribute pMailItem.VB_VarHelpID = -1

'----------------------------------------------------------
Public Property Get OutlookApp() As Outlook.Application
    Set OutlookApp = pOutlookApp
End Property

Public Property Let OutlookApp(ByVal vNewValue As Outlook.Application)
    Set pOutlookApp = vNewValue
End Property

'----------------------------------------------------------
Public Property Get MailItem() As Outlook.MailItem
    Set MailItem = pMailItem
End Property

Public Property Let MailItem(ByVal vNewValue As Outlook.MailItem)
    Set pMailItem = vNewValue
End Property

'----------------------------------------------------------
Public Property Get ReturnMail() As ClsMailItem
    Set ReturnMail = pReturnMail
End Property

Public Property Let ReturnMail(ByVal vNewValue As ClsMailItem)
    Set pReturnMail = vNewValue
End Property

'===============================================================
'Method DisplayEmail
'displays email before sending
'---------------------------------------------------------------
Public Sub DisplayEmail()
   On Error Resume Next
    With pMailItem
        .Display
    End With
    On Error GoTo 0
End Sub


'===============================================================
'Method CreateNewEmail
' creates a new email object
'---------------------------------------------------------------
Public Sub CreateNewEmail()
    If pOutlookApp Is Nothing Then GetOutlook
    If Not OutlookRunning Then
        Set pOutlookApp = Nothing
        GetOutlook
    End If
    Set pMailItem = pOutlookApp.CreateItem(olMailItem)
End Sub


'===============================================================
'Method SendEmail
'Sends email without displaying
'---------------------------------------------------------------
Public Sub SendEmail()
   On Error Resume Next
    With pMailItem
        If SEND_EMAILS Then .Send
    End With
    On Error GoTo 0
End Sub

'===============================================================
'Method ReturnMailItem
'Returns final mail item after it has been sent
'---------------------------------------------------------------
Private Function ReturnMailItem()
    Dim i As Integer
    
    With pReturnMail
        .MailTo = pMailItem.To
'        .Body = pMailItem.Body
        .Subject = pMailItem.Subject
        .timeSent = pMailItem.ReceivedTime
        .MailFrom = Environ("Username")
    End With
End Function

'===============================================================
'===============================================================
'SaveEmail
' saves email in database when it arrives
'---------------------------------------------------------------
Public Sub SaveEmail(MailItem As MailItem, Step As ClsStep)
    Dim RstEmail As Recordset
    Dim LastNo As Integer
    
    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblInbox WHERE StepIndex = " & Step.StepIndex)

    With RstEmail
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        !StepIndex = Step.StepIndex
        !Subject = MailItem.Subject
'        !Body = MailItem.Body
'        !To = MailItem.To
'        !CC = MailItem.CC
'        !Attachment = MailItem.Attachments.Item(1).PathName
        !EntryID = MailItem.EntryID
        .Update

    End With
    Set RstEmail = Nothing

End Sub

'===============================================================
' FindEmail
' Finds and returns email from EntryID
'---------------------------------------------------------------
Public Sub FindEmail(StepIndex As Integer)
    Dim OLNameSpace As Namespace
    Dim RstEmail As Recordset
    Dim FldrInbox As Outlook.Folder
    Dim EntryID As String
    Dim StoreID As String

    Set OLNameSpace = pOutlookApp.GetNamespace("MAPI")
    Set FldrInbox = OLNameSpace.GetDefaultFolder(olFolderInbox)
    Set RstEmail = ModDatabase.SQLQuery("SELECT EntryID FROM TblInbox WHERE StepIndex = " & StepIndex)

    With RstEmail
        If .RecordCount > 0 Then
            EntryID = !EntryID
        End If
    End With
    
    StoreID = FldrInbox.StoreID
    
    If EntryID <> "" Then
        On Error Resume Next
        Set pMailItem = OLNameSpace.GetItemFromID(EntryID, StoreID)
        On Error GoTo -1
        
        If Not pMailItem Is Nothing Then
            DisplayEmail
        End If
    End If
    
    Set OLNameSpace = Nothing
    Set RstEmail = Nothing
    Set FldrInbox = Nothing
End Sub

'===============================================================
'Event pOutlookApp_NewMail
'Event when new emails arrives
'---------------------------------------------------------------
Private Sub pOutlookApp_NewMail()
    If pOutlookApp Is Nothing Then GetOutlook
'    CheckInbox
End Sub

'===============================================================
' GetOutlook
' Starts outlook class
'---------------------------------------------------------------
Private Sub GetOutlook()
    On Error Resume Next
    
    Set pOutlookApp = GetObject(Class:="Outlook.Application")
    
    Err.Clear

    If pOutlookApp Is Nothing Then
        Shell ("Outlook")
        Set pOutlookApp = GetObject(Class:="Outlook.Application")
        
        Err.Clear
        
        If pOutlookApp Is Nothing Then
        
            Set pOutlookApp = CreateObject(Class:="Outlook.Application")
        End If
    End If
    
    Set pMailItem = pOutlookApp.CreateItem(0)
    Set pReturnMail = New ClsMailItem
End Sub

'===============================================================
'Event pMailItem_Send
'saves email details when it has been sent
'---------------------------------------------------------------
Private Sub pMailItem_Send(Cancel As Boolean)
    FrmProjectWF.BringToFront
End Sub

'===============================================================
Private Sub Class_Initialize()
    GetOutlook
End Sub

Private Sub Class_Terminate()
    Set pOutlookApp = Nothing
    Set pMailItem = Nothing
    Set pReturnMail = Nothing
End Sub
'===============================================================




