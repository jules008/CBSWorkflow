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
Private pProjectNo As Integer
Private pWorkflowType As enWorkflowType
Private pLender As ClsLender
Private pName As String
Private pMember As String
Private pRAG As enRAG
Private pCurrentStep As String
Private pNextStep As String
Private pPrevStep As String
Private pAltStep As String
Private pActiveStep As ClsStep
Private pStatus As String
Private pFrmInputBox As FrmInputBox
Private pSteps As ClsSteps
Private pChatRoom As ClsChatRoom
Private pProgress As Single
Private pParent As ClsProject

Private WithEvents pProjWFForm As FrmWFProject
Attribute pProjWFForm.VB_VarHelpID = -1
Private WithEvents pWFNameForm As FrmSelectWorkFlow
Attribute pWFNameForm.VB_VarHelpID = -1

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
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get WorkflowType() As enWorkflowType
    WorkflowType = pWorkflowType
End Property

Public Property Let WorkflowType(ByVal vNewValue As enWorkflowType)
    pWorkflowType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Lender() As ClsLender
    Set Lender = pLender
End Property

Public Property Let Lender(ByRef vNewValue As ClsLender)
    Set pLender = vNewValue
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
    pRAG = pActiveStep.RAG
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
Public Property Get AltStep() As String
    pAltStep = ActiveStep.AltStep
    AltStep = pAltStep
End Property

Public Property Let AltStep(ByVal vNewValue As String)
    pAltStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ActiveStep() As ClsStep
    Set pActiveStep = pSteps(pWorkflowNo & "-" & pCurrentStep)
    Set ActiveStep = pActiveStep
End Property

Public Property Let ActiveStep(ByRef vNewValue As ClsStep)
    vNewValue.Initialize Me
    Set pSteps(pWorkflowNo & "-" & pCurrentStep) = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As String
    Dim Step As ClsStep
    Dim Completed As Integer
    Dim NotStarted As Integer
    Dim i As Integer

    For i = 1 To pSteps.Count
        Set Step = pSteps(i)
        If Step.Status = "Complete" Then Completed = Completed + 1
        If Step.Status = "Not Started" Then NotStarted = NotStarted + 1
    Next

    If NotStarted = pSteps.Count Then
        pStatus = "Not Started"
    ElseIf Completed = pSteps.Count Then
        pStatus = "Complete"
    ElseIf ActiveStep.Status = "Waiting" Then
        pStatus = "Waiting"
    Else
        pStatus = "Action Req."
    End If
    
    Status = pStatus
    
    Set Step = Nothing
End Property

'---------------------------------------------------------------
Public Property Get Steps() As ClsSteps
    Set Steps = pSteps
End Property

Public Property Let Steps(ByRef vNewValue As ClsSteps)
    Set pSteps = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsProject
    Set Parent = pParent
End Property

Public Property Set Parent(newval As ClsProject)
    Set pParent = newval
End Property

'---------------------------------------------------------------
Public Property Get Progress() As Single
    Progress = pProgress
End Property

' ===============================================================
' Method DBNew
' displays new workflow form
' ---------------------------------------------------------------
Public Sub DBNew()
    pActiveStep.Start
    DBSave
    
    With pProjWFForm
        .Tag = "New"
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayWFSelectForm
' displays workflow select form
' ---------------------------------------------------------------
Public Sub DisplayWFSelectForm()
    With pWFNameForm
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    Dim ProgPC As Single
    Dim TmpWorkflow As ClsWorkflow
    Dim Adjust1 As Integer
    Dim Adjust2 As Integer
    
    ProgPC = pSteps.PCComplete
    
    With pProjWFForm
    
        If .ScreenAdjusted Then
            Adjust1 = 0
            Adjust2 = 0
        Else
            Adjust1 = 8
            Adjust2 = 16
            .ScreenAdjusted = True
        End If
        
        If pWorkflowType = enProject Then
            .TtlTop.Caption = "Project No " & Parent.ProjectNo & " Workflow"
            .LblLenderName.Visible = False
            .LblLenderCont.Visible = False
            .LblLenderPhone.Visible = False
            .TxtLenderContactName.Visible = False
            .TxtLenderContactPhone.Visible = False
            .TxtLenderName.Visible = False
            .xBtnChat.Visible = True
            .LblCl.Top = .LblCl.Top - Adjust2
            .LblCont1.Top = .LblCont1.Top - Adjust2
            .Lblphone1.Top = .Lblphone1.Top - Adjust2
            .TxtClientContactName.Top = .TxtClientContactName.Top - Adjust2
            .TxtClientContactPhone.Top = .TxtClientContactPhone.Top - Adjust2
            .TxtClientName.Top = .TxtClientName.Top - Adjust2
            .LblSPV.Top = .LblSPV.Top - Adjust1
            .LblCont2.Top = .LblCont2.Top - Adjust1
            .LblPhone2.Top = .LblPhone2.Top - Adjust1
            .TxtSPVContactName.Top = .TxtSPVContactName.Top - Adjust1
            .TxtSPVContactPhone.Top = .TxtSPVContactPhone.Top - Adjust1
            .TxtSPVName.Top = .TxtSPVName.Top - Adjust1
        Else
            .TtlTop.Caption = "Lender Workflow - Project No" & pProjectNo
            .TxtLenderName = pLender.Name
            .TxtLenderContactName = pLender.Contacts.PrimaryContact.ContactName
            .TxtLenderContactPhone = pLender.Contacts.PrimaryContact.Phone1
        End If
        
        .Progress ProgPC
        .TxtClientName = pParent.Client.Name
        .TxtSPVName = pParent.SPV.Name
        .TxtProjectNo = pParent.ProjectNo
        .TxtCaseManager = pParent.CaseManager.UserName
        .TxtLoanTerm = pParent.LoanTerm
        .TxtCommision = pParent.CBSComPC
        .ChkExitFee = pParent.ExitFee
        .TxtStepName = ActiveStep.StepNo & " - " & pActiveStep.StepName
        .xTxtAction = pActiveStep.StepAction
        .TxtClientContactName = Parent.Client.Contacts.PrimaryContact.ContactName
        .TxtClientContactPhone = Parent.Client.Contacts.PrimaryContact.Phone1
        .TxtSPVContactName = Parent.SPV.Contacts.PrimaryContact.ContactName
        .TxtSPVContactPhone = Parent.SPV.Contacts.PrimaryContact.Phone1
    
        Select Case pActiveStep.StepType
            Case enYesNo
                .TxtDataInput.Visible = False
                .BtnNo.Visible = True
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Yes"
            
            Case enStep
                .TxtDataInput.Visible = False
                .BtnNo.Visible = False
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Step Complete"
            
            Case enDataInput
                .TxtDataInput.Visible = True
                .TxtDataInput.Value = ""
                .BtnNo.Visible = False
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Step Complete"
       
                If pActiveStep.DataFormat = "Date" And .TxtDataInput = "" Then
                    Set TmpWorkflow = ActiveProject.ProjectWorkflow
                    .TxtDataInput = Format(FrmCalPicker.ShowForm, "dd mmm yy")
                    Set ActiveProject.ProjectWorkflow = TmpWorkflow
                End If
            
            Case enAltBranch
                .TxtDataInput.Visible = False
                .BtnCopyText.Visible = False
                .BtnNo.Visible = True
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Yes"
        End Select
        If Not .Visible Then .Show
    End With
    
    Set TmpWorkflow = Nothing
End Sub

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(WorkflowNo As String)
    Dim RstWorkflow As Recordset
    Dim IndexNo As Integer
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & WorkflowNo)
    With RstWorkflow
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!WorkflowNo) Then pWorkflowNo = !WorkflowNo
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!WorkflowType) Then pWorkflowType = enWorkflowTypeVal(!WorkflowType)
            If Not IsNull(!CurrentStep) Then pCurrentStep = !CurrentStep
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!RAG) Then pRAG = enRAGVal(!RAG)
            If Not IsNull(!Member) Then pMember = !Member
            If Not IsNull(!Progress) Then pProgress = !Progress
            
            pLender.DBGet !LenderNo
            Set pSteps = New ClsSteps
            pSteps.Initialize Me
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

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & pWorkflowNo)
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
        !CurrentStep = CleanSQLText(pCurrentStep, True)
        !WorkflowNo = CleanSQLText(pWorkflowNo, True)
        !ProjectNo = CleanSQLText(pProjectNo, True)
        !WorkflowType = CleanSQLText(enWorkflowTypeStr(pWorkflowType), True)
        !Name = CleanSQLText(pName, True)
        !Status = CleanSQLText(Status, True)
        !RAG = CleanSQLText(enRAGStr(RAG), True)
        !LenderNo = CleanSQLText(pLender.LenderNo, True)
        
        If Not Parent Is Nothing Then !ProjectNo = Parent.ProjectNo
        .Update
                    
        UpdateProgress
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

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo = " & pWorkflowNo)
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
    With ActiveStep
        .Complete
        .DBSave
    End With
    
        pCurrentStep = AltStep
        DBSave
    With ActiveStep
        .Start
            .DBSave
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
    Else
        With ActiveStep
            .Complete
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
    Dim InputForm As ClsInputBox
    
    Set InputForm = New ClsInputBox
    
    With InputForm
        .Title = "Pause Workflow"
        .Instructions = "How many days would you like to pause the workflow for?"
        .ClearForm
        .Show
        PauseDays = .ReturnValue
    End With
    
    Unload pFrmInputBox
    
    If IsNumeric(PauseDays) Then
        With ActiveStep
            .AmberThresh = .AmberThresh + PauseDays
            .RedThresh = .RedThresh + PauseDays
            .DBSave
        End With
    End If
    
    Set InputForm = Nothing
End Sub

' ===============================================================
' Method UpdateProgress
' Calculates progress of workflow, including lender workflows if it is type project
' ---------------------------------------------------------------
Private Sub UpdateProgress()
    
End Sub

' ===============================================================
' Event pWFNameForm_Update
' Event on selection of workflow type
' ---------------------------------------------------------------
Private Sub pWFNameForm_Update()
    With pWFNameForm
        pName = .GetWFName
    
    End With
End Sub

' ===============================================================
' Event pProjWFForm_ClickNo
' Event on clicking no button
' ---------------------------------------------------------------
Private Sub pProjWFForm_ClickNo()
    pProjWFForm.Hide
    MoveToAltStep
    With pProjWFForm
        If pActiveStep.LastStep Then
            Unload pProjWFForm
        Else
            If pActiveStep.Wait Then
                DBSave
                .Hide
            Else
                DisplayForm
            End If
        End If
    End With
End Sub

' ===============================================================
' Event pProjWFForm_CloseForm
' Event on clicking close workflow
' ---------------------------------------------------------------
Private Sub pProjWFForm_CloseForm()
    DBSave
    Unload pProjWFForm
End Sub

' ===============================================================
' Event pProjWFForm_PrevStep
' Event on clicking previous step
' ---------------------------------------------------------------
Private Sub pProjWFForm_PrevStep()
    pProjWFForm.Hide
    MoveToPrevStep
    pActiveStep.Start
    DBSave
    
    pProjWFForm.TxtDataInput = ""
    
    DisplayForm
    
End Sub

' ===============================================================
' Event pProjWFForm_StepComplete
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pProjWFForm_StepComplete()
    pProjWFForm.Hide
    With pProjWFForm
        If .TxtDataInput <> "" Then
            pActiveStep.DataItem = .TxtDataInput
            .TxtDataInput = ""
        End If
        
        MoveToNextStep
        
        If pActiveStep.LastStep Then
            Unload pProjWFForm
        Else
            If pActiveStep.Wait = True Then
                DBSave
                pProjWFForm.Hide
            Else
                DisplayForm
'                .Visible = True
            End If
        End If
    End With
End Sub

' ===============================================================
' Event pProjWFForm_StartChat
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pProjWFForm_StartChat()
    pChatRoom.DisplayForm pParent.ProjectNo
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsProject)
    Set Me.Parent = obj
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pSteps = New ClsSteps
    Set pActiveStep = New ClsStep
    Set pFrmInputBox = New FrmInputBox
    Set pLender = New ClsLender
    Set pProjWFForm = New FrmWFProject
    Set pWFNameForm = New FrmSelectWorkFlow
    Set pChatRoom = New ClsChatRoom
    
    pActiveStep.Initialize Me
    pSteps.Initialize Me
End Sub

' ===============================================================
Public Sub Terminate()
    pSteps.Terminate
    pActiveStep.Terminate
    
    Set pParent = Nothing
    Set pSteps = Nothing
    Set pActiveStep = Nothing
    Set pFrmInputBox = Nothing
    Set pLender = Nothing
    Set pProjWFForm = Nothing
    Set pWFNameForm = Nothing
    Set pChatRoom = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pParent = Nothing
    Set pSteps = Nothing
    Set pActiveStep = Nothing
    Set pFrmInputBox = Nothing
    Set pLender = Nothing
    Set pProjWFForm = Nothing
    Set pWFNameForm = Nothing
    Set pChatRoom = Nothing
End Sub
'---------------------------------------------------------------
