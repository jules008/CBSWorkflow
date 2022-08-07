VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsWorkflow
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Jun 20
'===============================================================

Option Explicit
Private pDeleted As Date
Private pWorkflowNo As Integer
Private pName As String
Private pMember As String
Private pRAG As enRAG
Private pCurrentStep As String
Private pNextStep As String
Private pPrevStep As String
Private pActiveStep As ClsStep
Private pStatus As enStatus
Private pFrmInputBox As FrmInputBox
Private pSteps As ClsSteps

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get WorkflowNo() As Integer
    WorkflowNo = pWorkflowNo
End Property

Public Property Let WorkflowNo(ByVal vNewValue As Integer)
    pWorkflowNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Member() As String
    Member = pMember
End Property

Public Property Let Member(ByVal vNewValue As String)
    pMember = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RAG() As enRAG
    pRAG = ActiveStep.RAG
    RAG = pRAG
End Property

'---------------------------------------------------------------
Public Property Get CurrentStep() As String
    CurrentStep = pCurrentStep
End Property

Public Property Let CurrentStep(ByVal vNewValue As String)
    pCurrentStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NextStep() As String
    pNextStep = ActiveStep.NextStep
    NextStep = pNextStep
End Property

Public Property Let NextStep(ByVal vNewValue As String)
    pNextStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PrevStep() As String
    pPrevStep = ActiveStep.PrevStep
    PrevStep = pPrevStep
End Property

Public Property Let PrevStep(ByVal vNewValue As String)
    pPrevStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ActiveStep() As ClsStep
    Set ActiveStep = pSteps(pWorkflowNo & "-" & pCurrentStep)
End Property

Public Property Let ActiveStep(ByVal vNewValue As ClsStep)
    vNewValue.Initialize Me
    Set pSteps(pWorkflowNo & "-" & pCurrentStep) = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As enStatus
    Dim Step As ClsStep
    Dim Completed As Integer
    Dim NotStarted As Integer
    Dim i As Integer

    For i = 1 To pSteps.Count
        Set Step = pSteps(i)
        If Step.Status = enComplete Then Completed = Completed + 1
        If Step.Status = enNotStarted Then NotStarted = NotStarted + 1
    Next

    If NotStarted = pSteps.Count Then
        pStatus = enNotStarted
    ElseIf Completed = pSteps.Count Then
        pStatus = enComplete
    ElseIf ActiveStep.Status = enWaiting Then
        pStatus = enWaiting
    Else
        pStatus = enActionReqd
    End If
    
    Status = pStatus
    
    Set Step = Nothing
End Property

'---------------------------------------------------------------
Public Property Get Steps() As ClsSteps
    Set Steps = pSteps
End Property

Public Property Let Steps(ByVal vNewValue As ClsSteps)
    Set pSteps = vNewValue
End Property

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(WorkflowNo As String)
    Dim RstWorkflow As Recordset
    Dim IndexNo As Integer
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & WorkflowNo & " AND Deleted IS NULL")
    With RstWorkflow
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!WorkflowNo) Then pWorkflowNo = !WorkflowNo
            If Not IsNull(!CurrentStep) Then pCurrentStep = !CurrentStep
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!Status) Then pStatus = enStatusVal(!Status)
            If Not IsNull(!RAG) Then pRAG = enRAGVal(!RAG)
            If Not IsNull(!Member) Then pMember = !Member
            
            pSteps.DeleteCollection
            pSteps.GetCollection
        End If
    End With
Set RstWorkflow = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstWorkflow As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & pWorkflowNo & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(WorkflowNo) FROM TblWorkflow ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstWorkflow
        If .RecordCount = 0 Then
            .AddNew
            pWorkflowNo = LastNo + 1
        Else
            .Edit
        End If
        
        If pSteps.Count = 0 Then pSteps.GetCollection
                
        If CurrentStep = "" Then pCurrentStep = pSteps.Start
        !CurrentStep = pCurrentStep
        !WorkflowNo = pWorkflowNo
        !Name = pName
        !Status = enStatusStr(Status)
        !RAG = enRAGStr(RAG)
        .Update
                    
        Steps.SaveCollection
        
    End With
    
    Set RstWorkflow = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstWorkflow As Recordset
    Dim i As Integer

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & pWorkflowNo & " AND Deleted IS NULL")
    With RstWorkflow
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

    Set RstWorkflow = Nothing
End Sub

' ===============================================================
' Method MoveToPrevStep
' Clear current step and moves to previous
' ---------------------------------------------------------------
Public Sub MoveToPrevStep()
    
    If PrevStep <> "" Then
        With ActiveStep
            .Reset
            .DBSave
        End With
        
        If PrevStep <> "0" Then
            
            pCurrentStep = PrevStep
            DBSave
            With ActiveStep
                .Reset
                .DBSave
            End With
        End If
    End If
End Sub

' ===============================================================
' Method MoveToAltStep
' Clear current step and moves to Alt step
' ---------------------------------------------------------------
Public Sub MoveToAltStep()
    Dim AltStep As String
    
    With ActiveStep
        
'        If .AltEmail.EmailNo <> 0 Then .AltEmail.Display
        
        Select Case .StepType
            
            Case enAltBranch

                AltStep = ActiveStep.AltStep
                
                If AltStep <> "" Then
'                    Do While pCurrentStep <> AltStep And pCurrentStep <> ActiveStep.Parent.Steps.Start
                        MoveToPrevStep
'                    Loop
                End If
            
        End Select
    End With
    
    With ActiveStep
        .Start
    End With
End Sub

' ===============================================================
' Method MoveToNextStep
' Completes current step and starts the next
' ---------------------------------------------------------------
Public Sub MoveToNextStep()
    
        With ActiveStep
            .Complete
            .DBSave
        End With
        
        If NextStep <> "0" Then
            pCurrentStep = NextStep
            DBSave
            With ActiveStep
                .Start
                .DBSave
            End With
    End If
End Sub

' ===============================================================
' Method Reset
' Resets workflow to start
' ---------------------------------------------------------------
Public Sub Reset()
    Dim Step As ClsStep
    
    pCurrentStep = pSteps.Start
    pNextStep = pActiveStep.NextStep
    pPrevStep = ""
    
    For Each Step In Steps
        With Step
            .StartTime = 0
            .EndTime = 0
            .DBSave
        End With
    Next
    ActiveStep.Start
End Sub

' ===============================================================
' Method Pause
' Pause workflow
' ---------------------------------------------------------------
Public Sub Pause()
    Dim PauseDays As Integer
    
    Load pFrmInputBox
    
    With pFrmInputBox
        .Caption = "Pause Workflow"
        .LblMessage = "How many days would you like to pause the workflow for?"
        .TxtInput = ""
        .Show
            
        PauseDays = .TxtInput
    End With
    
    Unload pFrmInputBox
    
    If IsNumeric(PauseDays) Then
        With ActiveStep
            .AmberThresh = .AmberThresh + PauseDays
            .RedThresh = .RedThresh + PauseDays
            .DBSave
        End With
    End If
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pSteps = New ClsSteps
    Set pActiveStep = New ClsStep
    Set pFrmInputBox = New FrmInputBox
    pActiveStep.Initialize Me
    pSteps.Initialize Me
End Sub

' ===============================================================
Public Sub Terminate()
    pSteps.Terminate
    pActiveStep.Terminate
    
    Set pSteps = Nothing
    Set pActiveStep = Nothing
    Set pFrmInputBox = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pSteps = Nothing
    Set pActiveStep = Nothing
    Set pFrmInputBox = Nothing
End Sub
'---------------------------------------------------------------
