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
Private pCaseManager As ClsCBSUser
Private pLoanTerm As String
Private pProjectWorkflow As ClsWorkflow
Private pCBSComPC As Single
Private pExitFee As Boolean
Private pWorkflows As ClsWorkflows
Private pContacts As ClsContacts
Private pSPVNo As Integer
Private pParent As ClsSPV
Private pProjectWFNo As Integer

'---------------------------------------------------------------
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
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
Public Property Get Parent() As ClsSPV
    If pParent Is Nothing Then
        Set pParent = New ClsSPV
        pParent.DBGet pSPVNo
    End If
    
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsSPV)
    Set pParent = obj
End Property
'---------------------------------------------------------------
' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ProjectNo As Integer)
    Dim RstProject As Recordset

    Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & ProjectNo)
    With RstProject
        If .RecordCount > 0 Then
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!LoanTerm) Then pLoanTerm = !LoanTerm
            If Not IsNull(!CBSComPC) Then pCBSComPC = !CBSComPC
            If Not IsNull(!SPVNo) Then pSPVNo = !SPVNo
            If Not IsNull(!ExitFee) Then pExitFee = !ExitFee
            
        pCaseManager.DBGet !CaseManager
        pContacts.GetCollection
        pProjectWorkflow.DBGet !ProjectWFNo
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

        If CStr(pProjectNo) = "" Then !ProjectNo = Null Else !ProjectNo = pProjectNo
        If CStr(pLoanTerm) = "" Then !LoanTerm = Null Else !LoanTerm = pLoanTerm
        If CStr(pCBSComPC) = "" Then !CBSComPC = Null Else !CBSComPC = pCBSComPC
        If CStr(pExitFee) = "" Then !ExitFee = Null Else !ExitFee = pExitFee
        If CStr(pSPVNo) = "" Then !SPVNo = Null Else !SPVNo = pSPVNo
        If pProjectWorkflow.WorkflowNo <> 0 Then !ProjectWFNo = pProjectWorkflow.WorkflowNo
        If pParent Is Nothing Then !SPVNo = Null Else !SPVNo = pParent.SPVNo
        If CStr(pCaseManager.CBSUserNo) = "" Then !CaseManager = Null Else !CaseManager = pCaseManager.CBSUserNo

        pWorkflows.SaveCollection
        pContacts.SaveCollection
        pProjectWorkflow.DBSave
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
Public Sub Initialize(obj As ClsSPV)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pWorkflows = New ClsWorkflows
    Set pProjectWorkflow = New ClsWorkflow
    Set pContacts = New ClsContacts
    Set pCaseManager = New ClsCBSUser

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
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
    Set pProjectWorkflow = Nothing
    Set pContacts = Nothing
    Set pCaseManager = Nothing
End Sub
' ===============================================================
