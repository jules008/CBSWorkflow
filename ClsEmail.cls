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
'Private pAttachment As ClsDocument
Private pParent As ClsStep

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
    Dim WorkflowNo As Integer
    Dim StepNo As String
    
'    WorkflowNo = Parent.Parent.WorkflowNo
'    StepNo = Parent.Parent.CurrentStep
    
    Subject = "||" & WorkflowNo & ":" & StepNo & "|| - " & pSubject
    
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

''---------------------------------------------------------------
'Public Property Get Attachment() As ClsDocument
'     Set Attachment = pAttachment
'End Property
'
'Public Property Let Attachment(ByVal vNewValue As ClsDocument)
'
'     Set pAttachment = vNewValue
'End Property

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
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(EmailNo As Integer)
    Dim RstEmail As Recordset

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & EmailNo & " AND Deleted IS NULL")
    With RstEmail
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!EmailNo) Then pEmailNo = !EmailNo
            If Not IsNull(!TemplateName) Then pTemplateName = !TemplateName
            If Not IsNull(!Subject) Then pSubject = !Subject
            If Not IsNull(!MailTo) Then pMailTo = !MailTo
            If Not IsNull(!CC) Then pCC = !CC
            If Not IsNull(!Body) Then pBody = !Body
        
'            If Not IsNull(!Attachment) Then
'                Set pAttachment = New ClsDocument
'                pAttachment.SetParent ObjPtr(Me)
'                pAttachment.DBGet !Attachment
'            Else
'                Set pAttachment = Nothing
'            End If
            
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

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & pEmailNo & " AND Deleted IS NULL")
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
        !MailTo = pMailTo
        !CC = pCC
        !EmailNo = pEmailNo
        !TemplateName = pTemplateName
        !Subject = pSubject
        !Body = pBody
        .Update

'        If Not pAttachment Is Nothing Then pAttachment.DBSave
        
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

    Set RstEmail = ModDatabase.SQLQuery("SELECT * FROM TblEmail WHERE EmailNo = " & pEmailNo & " AND Deleted IS NULL")
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
'    MailSystem.CreateNewEmail
    
'    FrmWorkflow.Hide
    
    Me.ProgressBar.Show
    
'    If pAttachment Is Nothing Then
'        Me.ProgressBar.Progress "Generating Email....", 4 / 7 * 100
'    Else
'        Me.ProgressBar.Progress "Generating Email....", 1 / 7 * 100
'    End If
    
    Me.ProgressBar.Progress "Generating Email....", 4 / 7 * 100
    
'    With MailSystem.MailItem
'
'        .To = ModWorkflow.ReplaceKeyWords(pMailTo, Parent.Parent)
'        .CC = ModWorkflow.ReplaceKeyWords(pCC, Parent.Parent)
'        .Body = ModWorkflow.ReplaceKeyWords(pBody, Parent.Parent)
'        .Subject = ModWorkflow.ReplaceKeyWords(Subject, Parent.Parent)
'
'        If Not pAttachment Is Nothing Then
'
'            Me.ProgressBar.Progress "Creating Attachment....", 2 / 7 * 100
'            If pAttachment.DocumentNo > 0 Then
'                pAttachment.PrepareDocument
'
'                Me.ProgressBar.Progress "Attaching Document to Email....", 7 / 7 * 100
'
'                .Attachments.Add GetDocLocalPath(ThisWorkbook.Path) & TMP_FILES & pAttachment.Name
'            End If
'        End If
        
        pProgressBar.Hide
        
'        .Display
        
'        FrmWorkflow.Refresh
        
'    End With
End Sub

' ===============================================================
' Method Send
' builds and displays email via mailsystem class
'---------------------------------------------------------------
Public Sub Send()
'    MailSystem.CreateNewEmail
    
'    With MailSystem.MailItem
'        .To = ModWorkflow.ReplaceKeyWords(pMailTo, Parent.Parent)
'        .CC = ModWorkflow.ReplaceKeyWords(pCC, Parent.Parent)
'        .Body = ModWorkflow.ReplaceKeyWords(pBody, Parent.Parent)
'        .Subject = ModWorkflow.ReplaceKeyWords(Subject, Parent.Parent)
''        .Attachments.Add GetDocLocalPath(ThisWorkbook.Path) & TMP_FILES & pAttachment.Name
'        .Send
'    End With
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsStep)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pProgressBar = New FrmProgressBar
End Sub
' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    Set pProgressBar = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pProgressBar = Nothing
    
'    If Not pAttachment Is Nothing Then
'        pAttachment.SetParent 0
'        Set pAttachment = Nothing
'    End If
End Sub

'---------------------------------------------------------------
