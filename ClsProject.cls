VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsProject
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 03 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - ProjectNo
'---------------------------------------------------------------
Option Explicit

Private pProjectNo As Integer
Private pProjectName As String
Private pCaseManager As ClsCBSUser
Private pLoanTerm As String
Private pProjectWorkflow As ClsWorkflow
Private pCBSComPC As Single
Private pExitFee As Boolean
Private pWorkflows As ClsWorkflows
Private pContacts As ClsContacts
Private pParent As ClsSPV
Private pProjectWFNo As Integer
Private pClient As ClsClient
Private pSPV As ClsSPV

Private WithEvents pProjectForm As FrmProjectForm
Attribute pProjectForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ProjectName() As String
    ProjectName = pProjectName
End Property

Public Property Let ProjectName(ByVal ProjectName As String)
    pProjectName = ProjectName
End Property

'---------------------------------------------------------------
Public Property Get CaseManager() As ClsCBSUser
    Set CaseManager = pCaseManager
End Property

Public Property Let CaseManager(ByVal vNewValue As ClsCBSUser)
    Set pCaseManager = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LoanTerm() As String
    LoanTerm = pLoanTerm
End Property

Public Property Let LoanTerm(ByVal vNewValue As String)
    pLoanTerm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ProjectWorkflow() As ClsWorkflow
    Set ProjectWorkflow = pProjectWorkflow
End Property

Public Property Let ProjectWorkflow(ByRef vNewValue As ClsWorkflow)
    vNewValue.Initialize Me
    pProjectWorkflow.Name = "Project"
    Set pProjectWorkflow = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CBSComPC() As Single
    CBSComPC = pCBSComPC
End Property

Public Property Let CBSComPC(ByVal vNewValue As Single)
    pCBSComPC = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ExitFee() As Boolean
    ExitFee = pExitFee
End Property

Public Property Let ExitFee(ByVal vNewValue As Boolean)
    pExitFee = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Workflows() As ClsWorkflows
    Set Workflows = pWorkflows
End Property

Public Property Let Workflows(ByRef vNewValue As ClsWorkflows)
    Set pWorkflows = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Contacts() As ClsContacts
    Set Contacts = pContacts
End Property

Public Property Let Contacts(ByRef vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Client() As ClsClient
    Set Client = pClient
End Property

Public Property Let Client(ByRef Client As ClsClient)
    Set pClient = Client
End Property

'---------------------------------------------------------------
Public Property Get SPV() As ClsSPV
    Set SPV = pSPV
End Property

Public Property Let SPV(ByRef SPV As ClsSPV)
    Set pSPV = SPV
End Property
    
'---------------------------------------------------------------
Public Property Get Parent() As ClsSPV
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsSPV)
    Set pParent = obj
End Property
'---------------------------------------------------------------

' ===============================================================
' Method DBNew
' displays new SPV form
' ---------------------------------------------------------------
Public Sub DBNew()
    
    With pProjectForm
        .Tag = "New"
        .ClearForm
        .BtnUpdate.Caption = "Create"
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pProjectForm
        .TxtCBSComPC = pCBSComPC
        .TxtLoanTerm = pLoanTerm
        .TxtProjectNo = pProjectNo
        .TxtProjName = pProjectName
        If pCaseManager.CBSUserNo > 0 Then .CmoCaseManager = pCaseManager.CBSUserNo
        If pClient.ClientNo > 0 Then .CmoClientNo = pClient.ClientNo
        If pSPV.SPVNo > 0 Then .CmoSPVNo = pSPV.SPVNo
        .ChkExitFee = pExitFee
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ByVal ProjectNameNo As Variant)
    Dim RstProject As Recordset

    If IsNumeric(ProjectNameNo) Then
        Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & ProjectNameNo)
    Else
        Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectName = '" & ProjectNameNo & "'")
    End If
    
    With RstProject
        If .RecordCount > 0 Then
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!LoanTerm) Then pLoanTerm = !LoanTerm
            If Not IsNull(!CBSComPC) Then pCBSComPC = !CBSComPC
            If Not IsNull(!ExitFee) Then pExitFee = !ExitFee
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!ProjectName) Then pProjectName = !ProjectName
            
            pSPV.DBGet !SPVNo
            pClient.DBGet !ClientNo
        pCaseManager.DBGet !CaseManager
            pContacts.GetCollection "Project", CleanSQLText(pProjectName)
        pProjectWorkflow.DBGet !ProjectWFNo
            pWorkflows.GetCollection
        End If
    End With
    Set RstProject = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstProject As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & pProjectNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ProjectNo) FROM TblProject ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstProject
        If .RecordCount = 0 Then
            .AddNew
            pProjectNo = LastNo + 1
        Else
            .Edit
        End If

        pSPV.DBSave
        pClient.DBSave
        pCaseManager.DBSave
        
        pWorkflows.SaveCollection
        pContacts.SaveCollection
        pProjectWorkflow.DBSave
        
        !ProjectNo = CleanSQLText(pProjectNo, True)
        !LoanTerm = CleanSQLText(pLoanTerm, True)
        !CBSComPC = CleanSQLText(pCBSComPC, True)
        !ExitFee = CleanSQLText(pExitFee, True)
        !ProjectNo = CleanSQLText(pProjectNo, True)
        !ProjectName = CleanSQLText(pProjectName, True)
        !ProjectWFNo = CleanSQLText(pProjectWorkflow.WorkflowNo, True)
        !CaseManager = CleanSQLText(pCaseManager.CBSUserNo, True)
        !SPVNo = CleanSQLText(pSPV.SPVNo, True)
        !ClientNo = CleanSQLText(pClient.ClientNo, True)

        pSPV.ClientNo = pClient.ClientNo

        .Update
    End With

    Set RstProject = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblProject WHERE ProjectNo = " & pProjectNo

End Sub

' ===============================================================
' Event pContactForm_Update
' Event from client form to update Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_Update()
    Set pSPV = New ClsSPV
    Set pClient = New ClsClient
    Set pCaseManager = New ClsCBSUser
    
    With pProjectForm
       pCBSComPC = .TxtCBSComPC
       pLoanTerm = .TxtLoanTerm
       pClient.DBGet .CmoClientNo.List(.CmoClientNo.ListIndex, 0)
       pSPV.DBGet .CmoSPVNo.List(.CmoSPVNo - 1, 0)
       pCaseManager.DBGet .CmoCaseManager.List(.CmoCaseManager - 1, 0)
       pExitFee = .ChkExitFee
       pProjectName = .TxtProjName
    End With
    DBSave
End Sub

' ===============================================================
' Event pContactForm_CreateNew
' Event from SPV form to create new Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_CreateNew()
    With pProjectForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pContactForm_Delete
' Deletes Contact
' ---------------------------------------------------------------
Private Sub pProjectForm_Delete()
    DBDelete
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsSPV)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pWorkflows = New ClsWorkflows
    Set pProjectWorkflow = New ClsWorkflow
    Set pContacts = New ClsContacts
    Set pCaseManager = New ClsCBSUser
    Set pProjectForm = New FrmProjectForm
    Set pSPV = New ClsSPV
    Set pClient = New ClsClient
    
    With pProjectWorkflow
        .Name = "Project"
        .WorkflowType = enProject
    End With
    
    Set pParent = New ClsSPV

    pWorkflows.Initialize Me
    pProjectWorkflow.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pWorkflows.Terminate
    pProjectWorkflow.Terminate

    Set Me.Parent = Nothing

    Set pWorkflows = Nothing
    Set pProjectWorkflow = Nothing
    Set pContacts = Nothing
    Set pCaseManager = Nothing
    Set pProjectForm = Nothing
    Set pSPV = Nothing
    Set pClient = Nothing
    
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
    Set pProjectWorkflow = Nothing
    Set pContacts = Nothing
    Set pCaseManager = Nothing
    Set pProjectForm = Nothing
End Sub
' ===============================================================
