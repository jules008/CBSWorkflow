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
Private pStatus As enStatus
Private pWait As Boolean
Private pEmail As ClsEmail
Private pAltEmail As ClsEmail
Private pUniqueID As Integer
Private pParent As Long

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
    
    If pStatus = enComplete Then
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
    
    If pStatus = enComplete Then
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
'    ModWorkflow.ProcessDataInput
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
Public Property Get Status() As enStatus
    If pStartTime = 0 Then
        pStatus = enNotStarted
    ElseIf pStartTime <> 0 And pEndTime <> 0 Then
        pStatus = enComplete
    ElseIf pWait Then
        pStatus = enWaiting
    Else
        pStatus = enActionReqd
    End If
    Status = pStatus
End Property

Public Property Let Status(ByVal vNewValue As enStatus)
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
Public Property Get AltEmail() As ClsEmail
    Set AltEmail = pAltEmail
End Property

Public Property Let AltEmail(ByVal vNewValue As ClsEmail)
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
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function
'---------------------------------------------------------------
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
            If Not IsNull(!Status) Then pStatus = enStatusVal(!Status)
            If Not IsNull(!Wait) Then pWait = !Wait
            If Not IsNull(!CopyText) Then pCopyText = !CopyText
            If Not IsNull(!DataItem) Then pDataItem = !DataItem
            If Not IsNull(!DataFormat) Then pDataFormat = !DataFormat
            If Not IsNull(!DataDest) Then pDataDest = !DataDest
            If Not IsNull(!UniqueID) Then pUniqueID = !UniqueID
            If Not IsNull(!Email) And !Email <> 0 Then pEmail.DBGet !Email
            If Not IsNull(!AltEmail) And !AltEmail <> 0 Then pAltEmail.DBGet !AltEmail
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
            If Not IsNull(!Email) And !Email <> 0 Then pEmail.DBGet !Email
            If Not IsNull(!AltEmail) And !AltEmail <> 0 Then pAltEmail.DBGet !AltEmail
            
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
    If pEmail.EmailNo <> 0 Then pEmail.Display
End Sub

' ===============================================================
' Method Complete
' completes the step
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
    
    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE StepIndex = " & StepIndex & " AND Deleted IS NULL")
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
        !Status = enStatusStr(pStatus)
        !Email = pEmail.EmailNo
        !AltEmail = pAltEmail.EmailNo
        !Wait = pWait
        !WorkflowNo = Parent.WorkflowNo
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
        !Email = pEmail.EmailNo
        !AltEmail = pAltEmail.EmailNo
        !Wait = pWait
        .Update

    End With
    Set RstStep = Nothing
    Set RstMaxNo = Nothing
End Sub
' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsWorkflow
    Dim tmp As ClsWorkflow

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstStep As Recordset
    Dim i As Integer

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE StepNo = " & pStepNo & " AND Deleted IS NULL")
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
Private Sub Class_Initialize()
    Set pEmail = New ClsEmail
    Set pAltEmail = New ClsEmail
    
    pEmail.SetParent ObjPtr(Me)
    pAltEmail.SetParent ObjPtr(Me)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pEmail = Nothing
    
    If Not pEmail Is Nothing Then pEmail.SetParent 0
    If Not pAltEmail Is Nothing Then pAltEmail.SetParent 0
End Sub

'---------------------------------------------------------------

