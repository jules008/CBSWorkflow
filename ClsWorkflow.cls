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
Private pDebt As Single
Private pCBSComm As Single
Private pExitFee As Single
Private pLoanTerm As Integer
Private pLoanType As String
Private pSecondTier As String
Private pParent As ClsProject

Private WithEvents pProjWFForm As FrmProjectWF
Attribute pProjWFForm.VB_VarHelpID = -1
Private WithEvents pLenderWFForm As FrmLenderWF
Attribute pLenderWFForm.VB_VarHelpID = -1
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
    Dim NoComplete As Integer
    Dim NoSteps As Integer
    Dim AryStepNo() As String
    Dim ProjectWFNo As Integer
    Dim RstLenders As Recordset
    Dim SumProgress As Single
    Dim CntProgress As Integer
    
    If pProjectNo > 0 Then
    AryStepNo = Split(pCurrentStep, ".")
    NoComplete = AryStepNo(1)
    NoSteps = pSteps.Count
    
        If NoSteps > 0 Then pProgress = NoComplete / NoSteps * 100
        DB.Execute "UPDATE TblWorkflow SET Progress = " & pProgress & " WHERE WorkflowNo = " & pWorkflowNo
        
        Set RstLenders = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE ProjectNo = " & pProjectNo)
            SumProgress = 0
            With RstLenders
                .MoveLast
                .MoveFirst
                CntProgress = .RecordCount
                Do While Not .EOF
                If !WorkflowType = "enProject" Then ProjectWFNo = !ProjectNo
                    SumProgress = SumProgress + !Progress
                    .MoveNext
                Loop
            End With
            
            If ProjectWFNo > 0 Then
                DB.Execute "UPDATE TblWorkflow SET Progress = " & SumProgress / CntProgress & " WHERE ProjectNo = " & ProjectWFNo & " And WorkflowType = 'enProject'"
            End If
    End If
    Set RstLenders = Nothing
    
'    Debug.Print "NoSteps"; NoSteps, "NoComplete"; NoComplete, "pProgress"; pProgress
    
    Progress = pProgress
End Property

'---------------------------------------------------------------
Public Property Get Debt() As String
    Debt = pDebt
End Property

Public Property Let Debt(ByVal vNewValue As String)
    pDebt = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CBSComm() As String
    CBSComm = pCBSComm
End Property

Public Property Let CBSComm(ByVal vNewValue As String)
    pCBSComm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ExitFee() As String
    ExitFee = pExitFee
End Property

Public Property Let ExitFee(ByVal vNewValue As String)
    pExitFee = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LoanTerm() As Integer
    LoanTerm = pLoanTerm
End Property

Public Property Let LoanTerm(ByVal vNewValue As Integer)
    pLoanTerm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LoanType() As String
    LoanType = pLoanType
End Property

Public Property Let LoanType(ByVal LoanType As String)
    pLoanType = LoanType
End Property

'---------------------------------------------------------------
Public Property Get SecondTier() As String
    SecondTier = pSecondTier
End Property

Public Property Let SecondTier(ByVal SecondTier As String)
    pSecondTier = SecondTier
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
' Method DisplayLenderForm
' displays Lender Workflow form
' ---------------------------------------------------------------
Public Sub DisplayLenderForm()
    Dim ProgPC As Single
    Dim TmpWorkflow As ClsWorkflow
    Dim Adjust1 As Integer
    Dim Adjust2 As Integer
    
    ProgPC = Progress
    With pLenderWFForm
            .TtlTop.Caption = "Lender Workflow - Project No" & pProjectNo
            .TxtLenderName = pLender.Name
            .TxtLenderContactName = pLender.Contacts.PrimaryContact.ContactName
            .TxtLenderContactPhone = pLender.Contacts.PrimaryContact.Phone1
        .TxtDebt = pDebt
        .TxtCBSComm = pCBSComm
        .TxtExitFee = pExitFee
        .TxtLoanTerm = pLoanTerm
        .Progress ProgPC
        .TxtClientName = pParent.Client.Name
        .TxtProjectNo = pParent.ProjectNo
        .TxtCaseManager = pParent.CaseManager.UserName
        .TxtLoanTerm = pParent.LoanTerm
        .TxtLoanType = pLoanType
        .TxtSecondTier = pSecondTier
        .TxtStepName = ActiveStep.StepNo & " - " & pActiveStep.StepName
        .xTxtAction = pActiveStep.StepAction
        .TxtClientContactName = Parent.Client.Contacts.PrimaryContact.ContactName
        .TxtClientContactPhone = Parent.Client.Contacts.PrimaryContact.Phone1

         Select Case pActiveStep.StepType
             Case enYesNo
                 .LblDataIput.Visible = False
                 .xTxtDataInput.Visible = False
                 .BtnNo.Visible = True
                 .BtnComplete.Visible = True
                 .BtnComplete.Caption = "Yes"
             
             Case enStep
                 .LblDataIput.Visible = False
                 .xTxtDataInput.Visible = False
                 .BtnNo.Visible = False
                 .BtnComplete.Visible = True
                 .BtnComplete.Caption = "Step Complete"
             
             Case enDataInput
                 .LblDataIput.Visible = True
                 .xTxtDataInput.Visible = True
                 .xTxtDataInput.Value = ""
                 .BtnNo.Visible = False
                 .BtnComplete.Visible = True
                 .BtnComplete.Caption = "Step Complete"
        
                 If pActiveStep.DataFormat = "Date" And .xTxtDataInput = "" Then
                     Set TmpWorkflow = ActiveProject.ProjectWorkflow
                     .xTxtDataInput = Format(FrmCalPicker.ShowForm, "dd mmm yy")
                     Set ActiveProject.ProjectWorkflow = TmpWorkflow
        End If
        
             Case enAltBranch
                 .LblDataIput.Visible = False
                 .xTxtDataInput.Visible = False
                 .BtnNo.Visible = True
                 .BtnComplete.Visible = True
                 .BtnComplete.Caption = "Yes"
            End Select
        If Not .Visible Then .Show
    End With
    
    Set TmpWorkflow = Nothing
End Sub

' ===============================================================
' Method DisplayProjectForm
' displays project Workflow form
' ---------------------------------------------------------------
Public Sub DisplayProjectForm()
    Dim ProgPC As Single
    Dim TmpWorkflow As ClsWorkflow
    Dim Adjust1 As Integer
    Dim Adjust2 As Integer
    
    ProgPC = Progress
    With pProjWFForm
        .TtlTop.Caption = "Project No " & Parent.ProjectNo & " Workflow"
        .xBtnChat.Visible = True
        .Progress ProgPC
        .TxtCBSCommission = pParent.CBSCommission
        .TxtCBSCommPC = pParent.CBSCommPC
        .TxtDebt = pParent.Debt
        .TxtExitFee = pParent.ExitFee
        .TxtExitFeePC = pParent.ExitFeePC
        .TxtLoanTerm = pParent.LoanTerm
        .TxtClientName = pParent.Client.Name
        .TxtSPVName = pParent.SPV.Name
        .TxtProjectNo = pParent.ProjectNo
        .TxtCaseManager = pParent.CaseManager.UserName
        .TxtStepName = ActiveStep.StepNo & " - " & pActiveStep.StepName
        .xTxtAction = pActiveStep.StepAction
        .TxtClientContactName = Parent.Client.Contacts.PrimaryContact.ContactName
        .TxtClientContactPhone = Parent.Client.Contacts.PrimaryContact.Phone1
        .TxtSPVContactName = Parent.SPV.Contacts.PrimaryContact.ContactName
        .TxtSPVContactPhone = Parent.SPV.Contacts.PrimaryContact.Phone1
        .TxtConsComenceDte = Parent.ConsComenceDte
        .TxtMSName = Parent.MS.ContactName
        .TxtMSPhone = Parent.MS.Phone1
        .TxtMSEmail = Parent.MS.EmailAddress
        .TxtValName = Parent.Valuer.ContactName
        .TxtValPhone = Parent.Valuer.Phone1
        .TxtValEmail = Parent.Valuer.EmailAddress
    
        Select Case pActiveStep.StepType
            Case enYesNo
                .LblDataIput.Visible = False
                .xTxtDataInput.Visible = False
                .BtnNo.Visible = True
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Yes"
            
            Case enStep
                .LblDataIput.Visible = False
                .xTxtDataInput.Visible = False
                .BtnNo.Visible = False
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Step Complete"
            
            Case enDataInput
                .LblDataIput.Visible = True
                .xTxtDataInput.Visible = True
                .xTxtDataInput.Value = ""
                .BtnNo.Visible = False
                .BtnComplete.Visible = True
                .BtnComplete.Caption = "Step Complete"
       
                If pActiveStep.DataFormat = "Date" And .xTxtDataInput = "" Then
                    Set TmpWorkflow = ActiveProject.ProjectWorkflow
                    .xTxtDataInput = Format(FrmCalPicker.ShowForm, "dd mmm yy")
                    Set ActiveProject.ProjectWorkflow = TmpWorkflow
                End If
            
            Case enAltBranch
                .LblDataIput.Visible = False
                .xTxtDataInput.Visible = False
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
            If Not IsNull(!Debt) Then pDebt = !Debt
            If Not IsNull(!CBSComm) Then pCBSComm = !CBSComm
            If Not IsNull(!ExitFee) Then pExitFee = !ExitFee
            If Not IsNull(!LoanTerm) Then pLoanTerm = !LoanTerm
            If Not IsNull(!LoanType) Then pLoanType = !LoanType
            If Not IsNull(!SecondTier) Then pSecondTier = !SecondTier
            
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
        !Progress = Progress
            !Debt = CleanSQLText(pDebt, True)
            !CBSComm = CleanSQLText(pCBSComm, True)
            !ExitFee = CleanSQLText(pExitFee, True)
            !LoanTerm = CleanSQLText(pLoanTerm, True)
        !LoanType = CleanSQLText(pLoanType, True)
        !SecondTier = CleanSQLText(pSecondTier, True)
        If Not Parent Is Nothing Then !ProjectNo = Parent.ProjectNo
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
    pProjWFForm.Progress Progress
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
    pProjWFForm.Progress Progress
End Sub

' ===============================================================
' Method MoveToNextStep
' Completes current step and starts the next
' ---------------------------------------------------------------
Public Sub MoveToNextStep()
    Dim PrevStep As String
    
        With ActiveStep
        PrevStep = pCurrentStep
        .Complete
            .DBSave
        End With
        
        If NextStep <> "0" Then
            pCurrentStep = NextStep
            DBSave
            With ActiveStep
                .Start
            .PrevStep = PrevStep
                .DBSave
            End With
    Else
        With ActiveStep
            .Complete
            .DBSave
        End With
        
        With pParent
            .CompleteDate = Format(Now, "dd mmmm yy")
            .DBSave
        End With
    End If
    pProjWFForm.Progress Progress
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
' Event pLenderWFForm_ClickNo
' Event on clicking no button
' ---------------------------------------------------------------
Private Sub pLenderWFForm_ClickNo()
    pLenderWFForm.Hide
    MoveToAltStep
    With pLenderWFForm
        If pActiveStep.LastStep Then
            Unload pLenderWFForm
        Else
            If pActiveStep.Wait Then
                DBSave
                .Hide
            Else
                DisplayProjectForm
            End If
        End If
    End With
End Sub

' ===============================================================
' Event pLenderWFForm_CloseForm
' Event on clicking close lender workflow
' ---------------------------------------------------------------
Private Sub pLenderWFForm_CloseForm()
    DBSave
    Unload pLenderWFForm
End Sub

' ===============================================================
' Event pLenderWFForm_DisplayContacts
' Event on clicking display project contacts
' ---------------------------------------------------------------
Private Sub pLenderWFForm_DisplayContacts()
    Parent.DisplayContactFrm
End Sub

' ===============================================================
' Event pLenderWFForm_PrevStep
' Event on clicking previous step
' ---------------------------------------------------------------
Private Sub pLenderWFForm_PrevStep()
    pLenderWFForm.Hide
    MoveToPrevStep
    pActiveStep.Start
    DBSave
    
    pLenderWFForm.xTxtDataInput = ""
    
    DisplayProjectForm
    
End Sub

' ===============================================================
' Event pLenderWFForm_StartChat
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pLenderWFForm_StartChat()
    pChatRoom.DisplayForm pParent.ProjectNo
End Sub

' ===============================================================
' Event pLenderWFForm_StepComplete
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pLenderWFForm_StepComplete()
    pLenderWFForm.Hide
    With pLenderWFForm
        If .xTxtDataInput <> "" Then
            pActiveStep.DataItem = .xTxtDataInput
            .xTxtDataInput = ""
        End If
        
        MoveToNextStep
        
        If pActiveStep.LastStep Then
            Unload pLenderWFForm
        Else
            If pActiveStep.Wait = True Then
                DBSave
                pLenderWFForm.Hide
            Else
                DisplayLenderForm
'                .Visible = True
            End If
        End If
    End With
End Sub
    
' ===============================================================
' Event pLenderWFForm_UpdateLoan
' Updates Loan details
' ---------------------------------------------------------------
Private Sub pLenderWFForm_UpdateLoan()
    With pLenderWFForm
        pCBSComm = .TxtCBSComm
        pDebt = .TxtDebt
        pExitFee = .TxtExitFee
        pLoanTerm = .TxtLoanTerm
        DBSave
        MsgBox "Loan Details Updated", vbOKOnly + vbInformation, APP_NAME
    End With
End Sub

' ===============================================================
' Event pProjWFForm_DisplayContacts
' displays project contact form
' ---------------------------------------------------------------
Private Sub pProjWFForm_DisplayContacts()
    Parent.DisplayContactFrm
End Sub

' ===============================================================
' Event pProjWFForm_UpdateLoan
' Updates Loan details
' ---------------------------------------------------------------
Private Sub pProjWFForm_UpdateLoan()
    With pProjWFForm
    
        pParent.CBSCommission = .TxtCBSCommission
        pParent.CBSCommPC = .TxtCBSCommPC
        pParent.Debt = .TxtDebt
        pParent.ExitFee = .TxtExitFee
        pParent.ExitFeePC = .TxtExitFeePC
        pParent.LoanTerm = .TxtLoanTerm

        pParent.DBSave
        MsgBox "Loan Details Updated", vbOKOnly + vbInformation, APP_NAME
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
                DisplayProjectForm
            End If
        End If
    End With
End Sub

' ===============================================================
' Event pProjWFForm_CloseForm
' Event on clicking close lender workflow
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
    
    pProjWFForm.xTxtDataInput = ""
    
    DisplayProjectForm
    
End Sub

' ===============================================================
' Event pProjWFForm_StartChat
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pProjWFForm_StartChat()
    pChatRoom.DisplayForm pParent.ProjectNo
End Sub

' ===============================================================
' Event pProjWFForm_StepComplete
' Event on completion of workflow step
' ---------------------------------------------------------------
Private Sub pProjWFForm_StepComplete()
    pProjWFForm.Hide
    With pProjWFForm
        If .xTxtDataInput <> "" Then
            pActiveStep.DataItem = .xTxtDataInput
            .xTxtDataInput = ""
        End If
        
        MoveToNextStep
        
        If pActiveStep.LastStep Then
            Unload pProjWFForm
        Else
            If pActiveStep.Wait = True Then
                DBSave
                pProjWFForm.Hide
            Else
                DisplayProjectForm
'                .Visible = True
            End If
        End If
    End With
End Sub

' ===============================================================
' Event pWFNameForm_CloseFrm
' Event on form close
' ---------------------------------------------------------------
Private Sub pWFNameForm_CloseFrm()
    With pWFNameForm
        pName = ""
    End With
End Sub

' ===============================================================
' Event pWFNameForm_Update
' Event on selection of workflow type
' ---------------------------------------------------------------
Private Sub pWFNameForm_Update()
    With pWFNameForm
        pLoanType = .CmoLoanType
        pSecondTier = .CmoSecondTier
        pName = .CmoSecondTier
    End With
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
    Set pProjWFForm = New FrmProjectWF
    Set pLenderWFForm = New FrmLenderWF
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
    Set pLenderWFForm = Nothing
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
    Set pLenderWFForm = Nothing
    Set pWFNameForm = Nothing
    Set pChatRoom = Nothing
End Sub
'---------------------------------------------------------------
