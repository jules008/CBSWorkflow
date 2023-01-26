VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsStep
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
Private pStepIndex As Integer
Private pStepNo As String
Private pStepType As enStepType
Private pStepName As String
Private pStartTime As Date
Private pEndTime As Date
Private pStepAction As String
Private pNextStep As String
Private pPrevStep  As String
Private pFirstStep As Boolean
Private pLastStep As Boolean
Private pAltStep As String
Private pRAG As enRAG
Private pRedThresh As Integer
Private pAmberThresh As Integer
Private pTimeToAmber As Variant
Private pTimeToRed  As Variant
Private pAmberTime As Date
Private pRedTime As Date
Private pCopyTextName As String
Private pCopyText As String
Private pDataItem As Variant
Private pDataFormat As String
Private pDataDest As String
Private pWFName As String
Private pStatus As String
Private pWait As Boolean
Private pEmailNo As Integer
Private pEmail As ClsEmail
Private pUniqueID As Integer
Private pStepHelp As FrmStepHelp
Private pStepHelpTitle As String
Private pStepHelpBody As String
Private pParent As ClsWorkflow
Private WithEvents pStepForm As FrmAdnWFStep
Attribute pStepForm.VB_VarHelpID = -1
'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StepIndex() As Integer
    StepIndex = pStepIndex
End Property

Public Property Let StepIndex(ByVal vNewValue As Integer)
    pStepIndex = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StepNo() As String
    StepNo = pStepNo
End Property

Public Property Let StepNo(ByVal vNewValue As String)
    pStepNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StepType() As enStepType
    StepType = pStepType
End Property

Public Property Let StepType(ByVal vNewValue As enStepType)
    pStepType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StepName() As String
    StepName = pStepName
End Property

Public Property Let StepName(ByVal vNewValue As String)
    pStepName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StartTime() As Date
    StartTime = pStartTime
End Property

Public Property Let StartTime(ByVal vNewValue As Date)
    pStartTime = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EndTime() As Date
    EndTime = pEndTime
End Property

Public Property Let EndTime(ByVal vNewValue As Date)
    pEndTime = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StepAction() As String
    StepAction = pStepAction
End Property

Public Property Let StepAction(ByVal vNewValue As String)
    pStepAction = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NextStep() As String
    NextStep = pNextStep
End Property

Public Property Let NextStep(ByVal vNewValue As String)
    pNextStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PrevStep() As String
    PrevStep = pPrevStep
End Property

Public Property Let PrevStep(ByVal vNewValue As String)
    pPrevStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get FirstStep() As Boolean
    
    If pPrevStep = 0 Then
        pFirstStep = True
    Else
        pFirstStep = False
    End If
    
    FirstStep = pFirstStep
End Property

'---------------------------------------------------------------
Public Property Get LastStep() As Boolean
    
    If pNextStep = 0 Then
        pLastStep = True
    Else
        pLastStep = False
    End If
    
    LastStep = pLastStep
End Property

'---------------------------------------------------------------
Public Property Get AltStep() As String
    AltStep = pAltStep
End Property

Public Property Let AltStep(ByVal vNewValue As String)
    pAltStep = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RAG() As enRAG
    If DateDiff("s", Now, AmberTime) >= 0 Then
        pRAG = en3Green
    ElseIf DateDiff("s", Now, RedTime) >= 0 Then
        pRAG = en2Amber
    Else
        pRAG = en1Red
    End If
    RAG = pRAG
End Property

'---------------------------------------------------------------
Public Property Get RedThresh() As Integer
    RedThresh = pRedThresh
End Property

Public Property Let RedThresh(ByVal vNewValue As Integer)
    pRedThresh = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AmberThresh() As Integer
    AmberThresh = pAmberThresh
End Property

Public Property Let AmberThresh(ByVal vNewValue As Integer)
    pAmberThresh = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TimeToAmber() As Variant
    
    If pStatus = "Complete" Then
        TimeToAmber = AmberThresh
    ElseIf DateDiff("h", Now, AmberTime) > 0 Then
        pTimeToAmber = DateDiff("h", Now, AmberTime)
    Else
        pTimeToAmber = 0
    End If
    
    TimeToAmber = pTimeToAmber
End Property

'---------------------------------------------------------------
Public Property Get TimeToRed() As Variant
    
    If pStatus = "Complete" Then
        TimeToRed = RedThresh
    ElseIf DateDiff("h", Now, RedTime) > 0 Then
        pTimeToRed = DateDiff("h", Now, RedTime)
    Else
        pTimeToRed = 0
    End If
    
    TimeToRed = pTimeToRed
End Property

'---------------------------------------------------------------
Public Property Get AmberTime() As Date
    pAmberTime = DateAdd("d", pAmberThresh, pStartTime)
    AmberTime = pAmberTime
End Property

'---------------------------------------------------------------
Public Property Get RedTime() As Date
    pRedTime = DateAdd("d", pRedThresh, pStartTime)
    RedTime = pRedTime
End Property

'---------------------------------------------------------------
Public Property Get CopyTextName() As String
    CopyTextName = pCopyTextName
End Property

Public Property Let CopyTextName(ByVal vNewValue As String)
    pCopyTextName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Wait() As Boolean
    Wait = pWait
End Property

Public Property Let Wait(ByVal vNewValue As Boolean)
    pWait = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EmailNo() As Integer
    EmailNo = pEmailNo
End Property

Public Property Let EmailNo(ByVal vNewValue As Integer)
    pEmailNo = vNewValue
    If pEmailNo > 0 Then
        Set pEmail = New ClsEmail
        pEmail.Initialize Me
        pEmail.DBGet pEmailNo
    End If
End Property

'---------------------------------------------------------------
Public Property Get CopyText() As String
    CopyText = pCopyText
End Property

Public Property Let CopyText(ByVal vNewValue As String)
    pCopyText = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DataItem() As Variant
    DataItem = pDataItem
End Property

Public Property Let DataItem(ByVal vNewValue As Variant)
    pDataItem = vNewValue
    ModWorkflow.ProcessDataInput
End Property

'---------------------------------------------------------------
Public Property Get DataFormat() As String
    DataFormat = pDataFormat
End Property

Public Property Let DataFormat(ByVal vNewValue As String)
    pDataFormat = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DataDest() As String
    DataDest = pDataDest
End Property

Public Property Let DataDest(ByVal vNewValue As String)
    pDataDest = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get WFName() As String
    WFName = pWFName
End Property

Public Property Let WFName(ByVal vNewValue As String)
    pWFName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As String
    If pStartTime = 0 Then
        pStatus = "Not Started"
    ElseIf pStartTime <> 0 And pEndTime <> 0 Then
        pStatus = "Complete"
    ElseIf pWait Then
        pStatus = "Waiting"
    Else
        pStatus = "Action Req."
    End If
    Status = pStatus
End Property

Public Property Let Status(ByVal vNewValue As String)
    pStatus = vNewValue
End Property
'---------------------------------------------------------------
Public Property Get Email() As ClsEmail
    Set Email = pEmail
End Property

Public Property Let Email(ByVal vNewValue As ClsEmail)
    Set pEmail = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UniqueID() As Integer
    UniqueID = pUniqueID
End Property

Public Property Let UniqueID(ByVal vNewValue As Integer)
    pUniqueID = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsWorkflow
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsWorkflow)
    Set pParent = obj
End Property
'---------------------------------------------------------------

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    Set pStepForm = New FrmAdnWFStep
    With pStepForm
        .TxtAltStep = pAltStep
        .TxtAmberThresh = pAmberThresh
        .TxtDataDest = pDataDest
        .TxtNextStep = pNextStep
        .TxtRedThresh = pRedThresh
        .xTxtStepAction = pStepAction
        .TxtStepName = pStepName
        .TxtStepNo = pStepNo
        .TxtWFName = pWFName
        
        .ChkWait = pWait
        .CmoDataFormat = pDataFormat
        If pEmailNo > 0 Then .CmoEmail = pEmailNo
        .CmoStepType.ListIndex = pStepType

        .Show
    End With
    Set pStepForm = Nothing
End Sub

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(WorkflowNo As Integer, StepNo As String)
    Dim RstStep As Recordset

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE StepNo = '" & StepNo _
                                    & "' AND WorkflowNo = " & WorkflowNo)
    With RstStep
        If .RecordCount > 0 Then
            If Not IsNull(!StepIndex) Then pStepIndex = !StepIndex
            If Not IsNull(!StepNo) Then pStepNo = !StepNo
            If Not IsNull(!StepType) Then pStepType = enStepTypeVal(!StepType)
            If Not IsNull(!StepName) Then pStepName = !StepName
            If Not IsNull(!StartTime) Then pStartTime = !StartTime
            If Not IsNull(!EndTime) Then pEndTime = !EndTime
            If Not IsNull(!StepAction) Then pStepAction = !StepAction
            If Not IsNull(!NextStep) Then pNextStep = !NextStep
            If Not IsNull(!PrevStep) Then pPrevStep = !PrevStep
            If Not IsNull(!AltStep) Then pAltStep = !AltStep
            If Not IsNull(!RedThresh) Then pRedThresh = !RedThresh
            If Not IsNull(!AmberThresh) Then pAmberThresh = !AmberThresh
            If Not IsNull(!CopyTextName) Then pCopyTextName = !CopyTextName
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!Wait) Then pWait = !Wait
            If Not IsNull(!EmailNo) Then EmailNo = !EmailNo
            If Not IsNull(!CopyText) Then pCopyText = !CopyText
            If Not IsNull(!DataItem) Then pDataItem = !DataItem
            If Not IsNull(!DataFormat) Then pDataFormat = !DataFormat
            If Not IsNull(!DataDest) Then pDataDest = !DataDest
            If Not IsNull(!UniqueID) Then pUniqueID = !UniqueID
            If Not IsNull(!HelpText) Then
                Dim HelpText() As String
                HelpText = Split(!HelpText, ":")
                pStepHelpTitle = HelpText(0)
                pStepHelpBody = HelpText(1)
            End If
            
        End If
    End With
Set RstStep = Nothing
End Sub

' ===============================================================
' Method DBGetTemplate
' Gets blank step templates for new classes
'---------------------------------------------------------------
Public Sub DBGetTemplate(StepNo As String)
    Dim RstStep As Recordset

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStepTemplate WHERE StepNo = '" & StepNo & "'")
    With RstStep
        If .RecordCount > 0 Then
            If Not IsNull(!StepNo) Then pStepNo = !StepNo
            If Not IsNull(!StepType) Then pStepType = enStepTypeVal(!StepType)
            If Not IsNull(!StepIndex) Then pStepIndex = !StepIndex
            If Not IsNull(!StepName) Then pStepName = !StepName
            If Not IsNull(!StepAction) Then pStepAction = !StepAction
            If Not IsNull(!NextStep) Then pNextStep = !NextStep
            If Not IsNull(!PrevStep) Then pPrevStep = !PrevStep
            If Not IsNull(!AltStep) Then pAltStep = !AltStep
            If Not IsNull(!RedThresh) Then pRedThresh = !RedThresh
            If Not IsNull(!AmberThresh) Then pAmberThresh = !AmberThresh
            If Not IsNull(!CopyTextName) Then pCopyTextName = !CopyTextName
            If Not IsNull(!CopyText) Then pCopyText = !CopyText
            If Not IsNull(!DataItem) Then pDataItem = !DataItem
            If Not IsNull(!DataFormat) Then pDataFormat = !DataFormat
            If Not IsNull(!DataDest) Then pDataDest = !DataDest
            If Not IsNull(!WFName) Then pWFName = !WFName
            If Not IsNull(!UniqueID) Then pUniqueID = !UniqueID
            If Not IsNull(!Wait) Then pWait = !Wait
            If Not IsNull(!EmailNo) Then EmailNo = !EmailNo
            If Not IsNull(!HelpText) Then
                Dim HelpText() As String
                HelpText = Split(!HelpText, ":")
                pStepHelpTitle = HelpText(0)
                pStepHelpBody = HelpText(1)
            End If
            
        End If
    End With
Set RstStep = Nothing
End Sub

' ===============================================================
' Method Start
' Starts the step
'---------------------------------------------------------------
Public Sub Start()
    
    pStartTime = Now
    DBSave
    If pEmailNo <> 0 Then pEmail.Display
    
End Sub

' ===============================================================
' Method Complete
' Completes the step
'---------------------------------------------------------------
Public Sub Complete()
    pEndTime = Now
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave(Optional StepIndex As String)
    Dim RstStep As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    If StepIndex = "" Then StepIndex = pStepIndex
    
    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE StepIndex = " & StepIndex)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(StepIndex) FROM TblStep ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstStep
        If .RecordCount = 0 Or StepIndex = 0 Then
            .AddNew
            pStepIndex = LastNo + 1
        Else
            .Edit
        End If
        
        !StepIndex = pStepIndex
        !StepNo = pStepNo
        !StepType = enStepTypeStr(pStepType)
        !StepName = pStepName
        !StartTime = pStartTime
        !EndTime = pEndTime
        !StepAction = pStepAction
        !NextStep = pNextStep
        !PrevStep = pPrevStep
        !AltStep = AltStep
        !RedThresh = pRedThresh
        !AmberThresh = pAmberThresh
        !CopyTextName = pCopyTextName
        !CopyText = pCopyText
        !DataItem = pDataItem
        !DataFormat = pDataFormat
        !DataDest = pDataDest
        !UniqueID = pUniqueID
        !Status = pStatus
        !EmailNo = pEmailNo
        !Wait = pWait
        !WorkflowNo = Parent.WorkflowNo
        !HelpText = pStepHelpTitle & ":" & pStepHelpBody
        .Update

    End With
    Set RstStep = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBSaveTemplate
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSaveTemplate(StepNo As String)
    Dim RstStep As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStepTemplate WHERE StepNo = '" & StepNo & "'")

    With RstStep
        .Edit
        !StepIndex = pStepIndex
        !StepNo = pStepNo
        !StepType = enStepTypeStr(pStepType)
        !StepName = pStepName
        !StepAction = pStepAction
        !NextStep = pNextStep
        !PrevStep = pPrevStep
        !AltStep = AltStep
        !RedThresh = pRedThresh
        !AmberThresh = pAmberThresh
        !CopyTextName = pCopyTextName
        !CopyText = pCopyText
        !DataItem = pDataItem
        !DataFormat = pDataFormat
        !DataDest = pDataDest
        !UniqueID = pUniqueID
        !EmailNo = pEmailNo
        !Wait = pWait
        .Update

    End With
    Set RstStep = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstStep As Recordset
    Dim i As Integer

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE StepNo = " & pStepNo)
    With RstStep
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

    Set RstStep = Nothing
End Sub

' ===============================================================
' Method Reset
' Resets step to beginning
'---------------------------------------------------------------
Public Sub Reset()
    pStartTime = 0
    pEndTime = 0
End Sub

' ===============================================================
' Method DisplayHelpForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayHelpForm()
    Set pStepHelp = New FrmStepHelp
    With pStepHelp
        .TxtTitle = pStepHelpTitle
        .xTxtHelpText = pStepHelpBody
        .Show
    End With
    Set pStepHelp = Nothing
End Sub

' ===============================================================
' Event pStepForm_Update
' Event from Step form to update step
' ---------------------------------------------------------------
Private Sub pStepForm_Update()
    With pStepForm
        pAltStep = .TxtAltStep
        pAmberThresh = .TxtAmberThresh
        pDataDest = .TxtDataDest
        pNextStep = .TxtNextStep
        pRedThresh = .TxtRedThresh
        pStepAction = .xTxtStepAction
        pStepName = .TxtStepName
        pStepNo = .TxtStepNo
        pStepType = .CmoStepType.ListIndex
        pWait = .ChkWait

        If .CmoEmail.ListIndex <> -1 Then
            pEmail.DBGet .CmoEmail.List(.CmoEmail.ListIndex, 0)
        End If
        
        If .CmoDataFormat.ListIndex <> -1 Then
            pDataFormat = .CmoDataFormat.List(.CmoDataFormat.ListIndex, 0)
        End If
        
        DBSaveTemplate pStepNo

    End With
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsWorkflow)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------
Private Sub Class_Initialize()
    
End Sub
' ===============================================================
Public Sub Terminate()
    
    If Not pEmail Is Nothing Then pEmail.Terminate

    Set pParent = Nothing
    Set pStepHelp = Nothing
    Set pStepForm = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pStepHelp = Nothing
    Set pEmail = Nothing
    Set pStepForm = Nothing
End Sub

'---------------------------------------------------------------

