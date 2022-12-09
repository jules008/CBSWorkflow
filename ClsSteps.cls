VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsSteps
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Jun 20
'===============================================================

Option Explicit
Private pSteps As Collection
Private pStart As String
Private pParent As ClsWorkflow

'---------------------------------------------------------------
Public Property Get Parent() As ClsWorkflow
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsWorkflow)
    Set pParent = obj
End Property

'---------------------------------------------------------------
Public Property Get Start() As String
    Start = pStart
End Property

Public Property Let Start(ByVal vNewValue As String)
    pStart = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pSteps.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(ByRef StepIndex As Variant) As ClsStep
Attribute FindItem.VB_UserMemId = 0
    
    On Error Resume Next
    Set FindItem = pSteps.Item(StepIndex)
    
    If FindItem Is Nothing Then
        Set FindItem = pSteps.Item(CStr(Parent.WorkflowNo & "-" & StepIndex))
    End If
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Step As ClsStep)
    Step.Initialize Me.Parent
    pSteps.Add Step, Key:=Parent.WorkflowNo & "-" & Step.StepNo
    
    If Step.PrevStep = "0" Then
        pStart = Step.StepNo
        
        If Parent.CurrentStep = "" Then
            Parent.CurrentStep = Step.StepNo
        End If
    End If
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(StepIndex As Variant)
    pSteps.Remove StepIndex
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pSteps.Count
End Function

' ===============================================================
' Method SaveCollection
' Saves the steps of a workflow
'---------------------------------------------------------------
Public Function SaveCollection()
    Dim Step As ClsStep
    
    Sin 0
    For Each Step In pSteps
        Step.DBSave
    Next
    Set Step = Nothing
    
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstStep As Recordset
    Dim Step As ClsStep
    Dim i As Integer

    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStep WHERE WorkflowNo = " & Parent.WorkflowNo _
                                        & " ORDER BY StepNo")
    
    If RstStep.RecordCount = 0 Then
        GetStepTemplates
    Else
        With RstStep
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Step = New ClsStep
                Step.DBGet Parent.WorkflowNo, !StepNo
                Me.AddItem Step
                .MoveNext
            Next
        End With
    End If
End Sub

' ===============================================================
' Method GetStepTemplates
' Gets a new set of blank step templates
'---------------------------------------------------------------
Public Sub GetStepTemplates()
    Dim Workflow As String
    Dim RstStep As Recordset
    Dim Step As ClsStep
    Dim i As Integer
    
    Workflow = Parent.Name
    
    Set RstStep = ModDatabase.SQLQuery("SELECT * FROM TblStepTemplate WHERE " _
                                        & "WFName = '" & Workflow & "' ORDER BY StepNo")
    
    If RstStep.RecordCount = 0 Then Exit Sub
    
    With RstStep
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Step = New ClsStep
            
            Step.DBGetTemplate !StepNo
            Me.AddItem Step
            Step.DBSave
            .MoveNext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Step As ClsStep
    Sin 0
    For Each Step In pSteps
       Set Step = Nothing
    Next
End Sub

' ===============================================================
' Method OpenNewEmails
' Opens any new emails in the steps
'---------------------------------------------------------------
'Public Sub OpenNewEmails()
'    Dim Step As ClsStep
'
'    For Each Step In pSteps
'        If Step.EmailArrived Then
'            MailSystem.FindEmail Step.StepIndex
'            Step.EmailArrived = False
'            Step.DBSave
'        End If
'    Next
'End Sub

' ===============================================================
' UpdateSteps
' Updates active steps with changes to step template
' ---------------------------------------------------------------
Public Sub UpdateSteps()
    Dim RstWFs As Recordset
    Dim RstSteps As Recordset
    Dim RstDoDCert As Recordset
    Dim RstTest As Recordset
    Dim RstNextRec As Recordset
    Dim SQL As String
    Dim WFNo As Integer
    Dim StepNo As String
    Dim DoDCertType As String
    Dim Response As Integer
    
    DB.Execute "UPDATE TblStep SET StartTime = NULL WHERE StartTime = 0"
    DB.Execute "UPDATE TblStep SET EndTime = NULL WHERE EndTime = 0"
    
    DB.Execute "UPDATE TblStep, TblStepTemplate" _
             & " SET" _
             & "    TblStep.Wait=TblStepTemplate.Wait," _
             & "    TblStep.StepType=TblStepTemplate.StepType," _
             & "    TblStep.StepNo=TblStepTemplate.StepNo," _
             & "    TblStep.StepName=TblStepTemplate.StepName," _
             & "    TblStep.StepAction=TblStepTemplate.StepAction," _
             & "    TblStep.NextStep=TblStepTemplate.NextStep," _
             & "    TblStep.PrevStep=TblStepTemplate.PrevStep," _
             & "    TblStep.AltStep=TblStepTemplate.AltStep," _
             & "    TblStep.RedThresh=TblStepTemplate.RedThresh," _
             & "    TblStep.AmberThresh=TblStepTemplate.AmberThresh," _
             & "    TblStep.DataItem=TblStepTemplate.DataItem," _
             & "    TblStep.DataFormat=TblStepTemplate.DataFormat," _
             & "    TblStep.DataDest=TblStepTemplate.DataDest," _
             & "    TblStep.Email=TblStepTemplate.Email," _
             & "    TblStep.AltEmail=TblStepTemplate.AltEmail," _
             & "    TblStep.CopyText=TblStepTemplate.CopyText," _
             & "    TblStep.CopyTextName = TblStepTemplate.CopyTextName" _
             & "  WHERE" _
             & "    TblStep.UniqueID = TblStepTemplate.UniqueID"
    
    SQL = "SELECT WorkflowNo FROM TblWorkflow WHERE Status <> 'Complete'"
    Set RstWFs = DB.OpenRecordset(SQL)
    
    With RstWFs
        Do While Not .EOF
        
            'cycle through WF Nos
            WFNo = !WorkflowNo
            'get cert type
            SQL = "SELECT Name FROM TblWorkflow WHERE WorkflowNo = " & WFNo & " AND Status <> 'Complete'"
            
            Set RstDoDCert = DB.OpenRecordset(SQL)
            
            If RstDoDCert.RecordCount <> 0 Then
            
                DoDCertType = RstDoDCert!Name
                
                SQL = "SELECT * FROM TblStepTemplate WHERE WFName = '" & DoDCertType & "'"
                Set RstSteps = DB.OpenRecordset(SQL)
                With RstSteps
                    'cycle through template steps
                    Do While Not .EOF
                        StepNo = !StepNo
                                            
                        'test to see whether record exists
                        SQL = "SELECT * FROM TblStep WHERE WorkflowNo = " & WFNo & " AND StepNo = '" & StepNo & "'"
                        Set RstTest = DB.OpenRecordset(SQL)
                        
                        With RstTest
                            If RstTest.RecordCount = 0 Then
                                .AddNew
                                !Wait = RstSteps!Wait
                                !StepType = RstSteps!StepType
                                !StepNo = RstSteps!StepNo
                                !StepName = RstSteps!StepName
                                !NextStep = RstSteps!NextStep
                                !StepAction = RstSteps!StepAction
                                !PrevStep = RstSteps!PrevStep
                                !AltStep = RstSteps!AltStep
                                !RedThresh = RstSteps!RedThresh
                                !AmberThresh = RstSteps!AmberThresh
                                !DataItem = RstSteps!DataItem
                                !DataFormat = RstSteps!DataFormat
                                !DataDest = RstSteps!DataDest
                                !Email = RstSteps!Email
                                !AltEmail = RstSteps!AltEmail
                                !WorkflowNo = WFNo
                                !Status = "Not Started"
                                !CopyText = RstSteps!CopyText
                                !CopyTextName = RstSteps!CopyTextName
                                !UniqueID = RstSteps!UniqueID
                                !StepIndex = !ID
                                
                                'Get Next record
                                SQL = "SELECT * FROM TblStep WHERE WorkflowNo = " & WFNo & " AND StepNo = '" & RstTest!NextStep & "'"
                                
                                Set RstNextRec = DB.OpenRecordset(SQL)
                                
                                If RstNextRec.RecordCount = 0 Then Err.Raise HANDLED_ERROR, , "Error loading WF " & WFNo
                                
                                'if the next record has been started, set record to be Complete
                                If RstNextRec!StartTime <> 0 And Not IsNull(RstNextRec!StartTime) Then
                                    !StartTime = RstNextRec!StartTime
                                    !EndTime = RstNextRec!StartTime
                                    !Status = "Complete"
                                    
                                End If
                                
                                .Update
                            End If
                        End With
                        .MoveNext
                    Loop
                End With
            End If
            
            UpdateCurrentStep WFNo
            
            .MoveNext
        Loop
    End With
    
    Set RstWFs = Nothing
    Set RstSteps = Nothing
    Set RstDoDCert = Nothing
    Set RstTest = Nothing
    Set RstNextRec = Nothing
End Sub

' ===============================================================
' UpdateCurrentStep
' Updates current step
' ---------------------------------------------------------------
Private Sub UpdateCurrentStep(WFNo As Integer)
    Dim RstWorkflows As Recordset
    Dim SQL As String
    Dim CurrentStep As String
    Dim Response As Integer
    
    SQL = "SELECT CurrentStep FROM TblWorkflow WHERE WorkflowNo = " & WFNo
    
    Set RstWorkflows = DB.OpenRecordset(SQL)
    
    With RstWorkflows
        If .RecordCount > 0 Then CurrentStep = !CurrentStep
    End With
    
    SQL = "SELECT StepNo FROM TblStep WHERE StartTime IS NOT NULL AND EndTime IS NULL AND WorkflowNo = " & WFNo
    
    Set RstWorkflows = DB.OpenRecordset(SQL)
    
    With RstWorkflows
        If .RecordCount = 0 Then Err.Raise HANDLED_ERROR, , "WF " & WFNo & " has no steps"
        .MoveLast
        If .RecordCount > 1 Then Err.Raise HANDLED_ERROR, , "WF " & WFNo & " has " & .RecordCount & " open steps"
        If .RecordCount = 1 Then
            If CurrentStep <> !StepNo Then
                DB.Execute "UPDATE TblWorkflow SET CurrentStep = '" & Format(!StepNo, "0.00") & "' WHERE WorkflowNo = " & WFNo
            End If
        End If
    End With
    
    Set RstWorkflows = Nothing
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsWorkflow)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pSteps = New Collection
End Sub
' ===============================================================
Public Sub Terminate()
    Dim Step As ClsStep

    Sin 0
    For Each Step In pSteps
        Step.Terminate
    Next
    Set pParent = Nothing
    Set pSteps = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pSteps = Nothing
End Sub

'---------------------------------------------------------------
