VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsWorkflows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsWorkflows
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 24 Jun 20
'===============================================================

Option Explicit
Private pWorkflows As Collection
Private pActive As Boolean
Private pDeleted As Boolean
'---------------------------------------------------------------

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pWorkflows.[_NewEnum]
End Property

'---------------------------------------------------------------
Public Property Get Active() As Boolean
    Active = pActive
End Property

Public Property Let Active(NewVal As Boolean)
    pActive = NewVal
End Property

'---------------------------------------------------------------
Public Property Get Deleted() As Boolean
    Deleted = pDeleted
End Property

Public Property Let Deleted(NewVal As Boolean)
    pDeleted = NewVal
End Property

' ===============================================================
' Method UpdateRAGs
' Updates the RAGs for all workflows
'---------------------------------------------------------------
Public Sub UpdateRAGs()
    Dim AmberDate As Date
    Dim RedDate As Date
    Dim RstWorkflow As Recordset
    Dim RAG As String
    Dim SQL As String
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT " _
                                                & "TblWorkflow.WorkflowNo, " _
                                                & "TblWorkflow.CurrentStep, " _
                                                & "TblStep.RedThresh, " _
                                                & "TblStep.AmberThresh, " _
                                                & "TblStep.StartTime, " _
                                                & "TblStep.EndTime, " _
                                                & "TblWorkflow.Status, " _
                                                & "TblWorkflow.RAG " _
                                            & "FROM " _
                                                & "TblWorkflow Inner Join " _
                                                & "TblStep On TblStep.WorkflowNo = TblWorkflow.WorkflowNo " _
                                            & "AND TblStep.StepNo = TblWorkflow.CurrentStep " _
                                            & "WHERE " _
                                                & "TblWorkflow.Status <> 'enComplete' " _
                                            & "ORDER BY " _
                                                & "TblWorkflow.WorkflowNo")
                            
    With RstWorkflow
        Do While Not .EOF
            If Not IsNull(!StartTime) Then
                AmberDate = DateAdd("d", !AmberThresh, !StartTime)
                RedDate = DateAdd("d", !RedThresh, !StartTime)
                
                If DateDiff("s", Now, AmberDate) >= 0 Then
                    RAG = "en3Green"
                ElseIf DateDiff("d", Now, RedDate) >= 0 Then
                    RAG = "en2Amber"
                Else
                    RAG = "en1Red"
                End If
                
                SQL = "UPDATE TblWorkflow SET RAG = '" & RAG & "' WHERE WorkflowNo = " & !WorkflowNo
                DB.Execute SQL
            End If
            .MoveNext
        Loop
    End With
    Set RstWorkflow = Nothing
End Sub

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------
Public Function FindItem(WorkflowNo As Variant) As ClsWorkflow
    On Error Resume Next
    Set FindItem = pWorkflows.Item(CStr(WorkflowNo))
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(Workflow As ClsWorkflow)
    On Error Resume Next
    pWorkflows.Add Workflow, Key:=CStr(Workflow.WorkflowNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(WorkflowNo As Variant)
    pWorkflows.Remove CStr(WorkflowNo)
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pWorkflows.Count
End Function

' ===============================================================
' Method GetActive
' Gets all active items in collection
'---------------------------------------------------------------
Public Sub GetActive()
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim i As Integer
    Dim Status As enStatus

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT " _
                                                & " TblWorkflow.* " _
                                            & "FROM " _
                                                & "TblMember " _
                                                & "INNER JOIN TblWorkflow ON TblMember.StudentID = TblWorkflow.Member " _
                                            & "WHERE " _
                                                & "TblWorkflow.Deleted IS NULL " _
                                            & "ORDER BY " _
                                                & "TblMember.DisplayName")

    If RstWorkflow.RecordCount = 0 Then Exit Sub
    
    With RstWorkflow
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Workflow = New ClsWorkflow
            Workflow.DBGet !WorkflowNo
            Me.AddItem Workflow
            .MoveNext
        Next
    End With
    pActive = True
End Sub

' ===============================================================
' Method GetDeleted
' Gets all deleted items in collection
'---------------------------------------------------------------
Public Sub GetDeleted()
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim i As Integer
    Dim Status As enStatus

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT " _
                                                & " TblWorkflow.* " _
                                            & "FROM " _
                                                & "TblMember " _
                                                & "INNER JOIN TblWorkflow ON TblMember.StudentID = TblWorkflow.Member " _
                                            & "WHERE " _
                                                & "TblWorkflow.Deleted IS NOT NULL " _
                                            & "ORDER BY " _
                                                & "TblMember.DisplayName")

    If RstWorkflow.RecordCount = 0 Then Exit Sub
    
    With RstWorkflow
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Workflow = New ClsWorkflow
            Workflow.DBGet !WorkflowNo
            Me.AddItem Workflow
            .MoveNext
        Next
    End With
    pDeleted = True
End Sub

' ===============================================================
' Method CountForAction
' returns number of workflows that require action
'---------------------------------------------------------------
Public Function CountForAction() As Integer
    Dim Workflow As ClsWorkflow
    Dim RstCount As Recordset
    Dim Count As Integer
    
    Set RstCount = ModDatabase.SQLQuery("SELECT COUNT(WorkflowNo) AS [CountX] " _
                                        & "FROM TblWorkflow " _
                                        & "WHERE ((RAG = 'en2Amber' Or RAG = 'en1Red') " _
                                        & "OR Status = 'enActionReqd') " _
                                        & "AND Deleted IS NULL " _
                                        & "AND Status <> 'enComplete'")
                                        
    CountForAction = RstCount![CountX]
    
    Set RstCount = Nothing
End Function

' ===============================================================
' Method GetComplete
' Gets all active items in collection
'---------------------------------------------------------------
Public Sub GetComplete()
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim i As Integer

    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE Status = 'enComplete' " _
                                            & " AND Deleted IS NULL")
    
    If RstWorkflow.RecordCount = 0 Then Exit Sub
    
    With RstWorkflow
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Workflow = New ClsWorkflow
            Workflow.DBGet !WorkflowNo
            Me.AddItem Workflow
            .MoveNext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Workflow As ClsWorkflow
    For Each Workflow In pWorkflows
        pWorkflows.Remove CStr(Workflow.WorkflowNo)
    Next
End Sub

' ===============================================================
' Method SaveCollection
' Deletes all items in collection
'---------------------------------------------------------------
Public Sub SaveCollection()
    Dim Workflow As ClsWorkflow
    
    For Each Workflow In pWorkflows
        Workflow.DBSave
    Next
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pWorkflows = New Collection

End Sub
' ===============================================================
Public Sub Terminate()
    Dim Workflow As ClsWorkflow

    For Each Workflow In pWorkflows
        Workflow.Terminate
    Next
    Set pWorkflows = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
End Sub

'---------------------------------------------------------------
