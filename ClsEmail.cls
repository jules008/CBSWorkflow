VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsEmail
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 01 Jul 20
'===============================================================

Option Explicit

Private pDeleted As Date
Private pEmailNo As Integer
Private pTemplateName As String
Private pSubject As String
Private pBody As String
Private pDateSent As Date
Private pMailTo As String
Private pCC As String
Private pProgressBar As FrmProgressBar
Private pParent As ClsStep
Private WithEvents pEmailForm As FrmAdnEmails
Attribute pEmailForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EmailNo() As Integer
    EmailNo = pEmailNo
End Property

Public Property Let EmailNo(ByVal vNewValue As Integer)
    pEmailNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TemplateName() As String
    TemplateName = pTemplateName
End Property

Public Property Let TemplateName(ByVal vNewValue As String)
    pTemplateName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Subject() As String
    Subject = pSubject
End Property
'---------------------------------------------------------------
Public Property Get SubjectText() As String
    SubjectText = pSubject
End Property

Public Property Let Subject(ByVal vNewValue As String)
    pSubject = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Body() As String
    Body = pBody
End Property

Public Property Let Body(ByVal vNewValue As String)
    pBody = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DateSent() As Date
    DateSent = pDateSent
End Property

Public Property Let DateSent(ByVal vNewValue As Date)
    pDateSent = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MailTo() As String
    MailTo = pMailTo
End Property

Public Property Let MailTo(ByVal vNewValue As String)
    pMailTo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CC() As String
    CC = pCC
End Property

Public Property Let CC(ByVal vNewValue As String)
    pCC = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ProgressBar() As FrmProgressBar
     Set ProgressBar = pProgressBar
End Property

Public Property Let ProgressBar(ByVal vNewValue As FrmProgressBar)
    
     Set pProgressBar = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsStep
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsStep)
    Set pParent = obj
End Property

' ===============================================================
' Method DBNew
' displays new User form
' ---------------------------------------------------------------
Public Sub DBNew()
    pCC = ""
    pEmailNo = 0
    pMailTo = ""
    pTemplateName = ""
    pBody = ""
    pSubject = ""
    
    With pEmailForm
        .Tag = "New"
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pEmailForm
        .Tag = "Update"
        .TxtCC = Trim(pCC)
        .TxtEmailNo = Trim(pEmailNo)
        .TxtMailTo = Trim(pMailTo)
        .TxtTemplateName = Trim(pTemplateName)
        .xTxtBody = Trim(pBody)
        .xTxtSubject = Trim(pSubject)
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(EmailNo As Integer)
    Dim RstEmail As Recordset

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & EmailNo)
    With RstEmail
        If .RecordCount > 0 Then
            If Not IsNull(!EmailNo) Then pEmailNo = !EmailNo
            If Not IsNull(!TemplateName) Then pTemplateName = !TemplateName
            If Not IsNull(!Subject) Then pSubject = !Subject
            If Not IsNull(!MailTo) Then pMailTo = !MailTo
            If Not IsNull(!CC) Then pCC = !CC
            If Not IsNull(!Body) Then pBody = !Body
        
            
        End If
    End With
Set RstEmail = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstEmail As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & pEmailNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(EmailNo) FROM TblEmail ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstEmail
        If .RecordCount = 0 Then
            .AddNew
            pEmailNo = LastNo + 1
        Else
            .Edit
        End If
        !MailTo = CleanSQLText(pMailTo, True)
        !CC = CleanSQLText(pCC, True)
        !EmailNo = CleanSQLText(pEmailNo, True)
        !TemplateName = CleanSQLText(pTemplateName, True)
        !Subject = CleanSQLText(pSubject, True)
        !Body = CleanSQLText(pBody, True)
        .Update

    End With
    Set RstEmail = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstEmail As Recordset
    Dim i As Integer

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & pEmailNo)
    With RstEmail
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
            Else
                .Edit
                !Deleted = Now
                .Update
            End If
        Next
    End With

    Set RstEmail = Nothing
End Sub

' ===============================================================
' Method Display
' builds and displays email via mailsystem class
'---------------------------------------------------------------
Public Sub Display()
    Dim MailItem As MailItem
    Dim MailAccount As Outlook.Account
    Dim OLPA As Outlook.PropertyAccessor
    Dim AttLogo As Outlook.Attachment
    Dim HTML1 As String
    Dim ImgPath As String
    
    ImgPath = GetDocLocalPath(ThisWorkbook.Path) & PICTURES_PATH & LOGO_FILE
    
    Set MailItem = Outlook.CreateItem(0)
    Set MailAccount = Session.Accounts("contact@cbs-capital.co.uk")
    Set AttLogo = MailItem.Attachments.Add(ImgPath)
    Set OLPA = AttLogo.PropertyAccessor
    
    OLPA.SetProperty PR_ATTACH_CONTENT_ID, "Logo"
    
    Me.ProgressBar.Show
    
    Me.ProgressBar.Progress "Generating Email....", 4 / 7 * 100
    HTML1 = ModWorkflow.ReplaceKeyWords(pBody, Parent.Parent)
    With MailItem
        .To = ModWorkflow.ReplaceKeyWords(pMailTo, Parent.Parent)
        .CC = ModWorkflow.ReplaceKeyWords(pCC, Parent.Parent)
        .Subject = ModWorkflow.ReplaceKeyWords(Subject, Parent.Parent)
        .BodyFormat = olFormatHTML
        .HTMLBody = ModWorkflow.ReplaceKeyWords(pBody, Parent.Parent)
        .SendUsingAccount = MailAccount
        .Display
        
    End With
        pProgressBar.Hide
End Sub

' ===============================================================
' Method Send
' builds and displays email via mailsystem class
'---------------------------------------------------------------
Public Sub Send()
    MailSystem.CreateNewEmail
    
    With MailSystem.MailItem
        .To = ModWorkflow.ReplaceKeyWords(pMailTo, Parent.Parent)
        .CC = ModWorkflow.ReplaceKeyWords(pCC, Parent.Parent)
        .Body = ModWorkflow.ReplaceKeyWords(pBody, Parent.Parent)
        .Subject = ModWorkflow.ReplaceKeyWords(Subject, Parent.Parent)
        .Send
    End With
End Sub


' ===============================================================
' Event pEmailForm_Update
' Event from client form to update CBS User
' ---------------------------------------------------------------
Private Sub pEmailForm_Update()

    With pEmailForm
        pCC = CleanSQLText(.TxtCC)
        pMailTo = CleanSQLText(.TxtMailTo)
        pTemplateName = CleanSQLText(.TxtTemplateName)
        pBody = CleanSQLText(.xTxtBody)
        pSubject = CleanSQLText(.xTxtSubject)
    End With
    DBSave
End Sub

' ===============================================================
' Event pEmailForm_CreateNew
' Event from user form to create new email
' ---------------------------------------------------------------
Private Sub pEmailForm_CreateNew()
    With pEmailForm
        .ClearForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pEmailForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pEmailForm_Delete()
    DBDelete
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsStep)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pProgressBar = New FrmProgressBar
    Set pEmailForm = New FrmAdnEmails
End Sub
' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    Set pProgressBar = Nothing
    Set pEmailForm = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pProgressBar = Nothing
    Set pEmailForm = Nothing
    
End Sub

'---------------------------------------------------------------
