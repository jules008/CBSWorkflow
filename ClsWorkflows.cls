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
' Date - 03 Oct 22
'===============================================================
' Methods
'---------------------------------------------------------------
' Find - Finds Item from index
' Add - Add item to Collection
' Remove - Removes item from collection
' Count - Returns number of items in collection
' GetCollection - Gets all items in collection
' SaveCollection - Saves all items in collection to database
' RemoveCollection - Removes all items in collection
' DeleteCollection - Deletes all items in collection
' Destroy - Removes item from collection and deletes it
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pWorkflows As Collection
Private pParent As ClsProject
Private WithEvents pWFTypeForm As FrmAdnWFTypes
Attribute pWFTypeForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get Parent() As ClsProject
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsProject)
    Set pParent = obj
End Property
'---------------------------------------------------------------
                
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pWorkflows.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(WorkflowNo As Variant) As ClsWorkflow
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pWorkflows.Item(WorkflowNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(ByRef Workflow As ClsWorkflow)
    Workflow.Initialize Me.Parent
    pWorkflows.Add Workflow, Key:=CStr(Workflow.WorkflowNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(WorkflowNo As Variant)
    pWorkflows(WorkflowNo).Parent = Nothing
    pWorkflows.Remove WorkflowNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pWorkflows.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim i As Integer
    
    Set RstWorkflow = ModDatabase.SQLQuery("TblWorkflow")
    With RstWorkflow
        If .RecordCount > 0 Then
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Workflow = New ClsWorkflow
            Workflow.DBGet !WorkflowNo
            Add Workflow
            .MoveNext
        Next
        End If
    End With
End Sub

' ===============================================================
' Method SaveCollection
' Saves all items in collection to database
' ---------------------------------------------------------------
Public Sub SaveCollection()
    Dim Workflow As ClsWorkflow
    
    Sin 0
    For Each Workflow In pWorkflows
        Workflow.DBSave
        Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Workflow As ClsWorkflow
    Sin 0
    For Each Workflow In pWorkflows
        pWorkflows.Remove CStr(Workflow.WorkflowNo)
        Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Workflow As ClsWorkflow
    Sin 0
    For Each Workflow In pWorkflows
        pWorkflows.Remove CStr(Workflow.WorkflowNo)
        Workflow.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(WorkflowNo As Variant)
    Dim Workflow As ClsWorkflow
    
    Set Workflow = pWorkflows(WorkflowNo)
    Workflow.Parent = Nothing
    Workflow.Terminate
    pWorkflows.Remove WorkflowNo
    Workflow.DBDelete True

    Set Workflow = Nothing
End Sub

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
' Method DisplayWFTypeFrm
' displays wf type form
'---------------------------------------------------------------
Public Sub DisplayWFTypeFrm(Optional WFNo As String)
    Dim RstWFType As Recordset
    
    If WFNo <> "" Then
        Set RstWFType = ModDatabase.SQLQuery("SELECT * FROM TblWorkflowType WHERE WFNo = " & WFNo)
        
        With RstWFType
            If .RecordCount > 0 Then
                pWFTypeForm.TxtDescription = !Description
                pWFTypeForm.TxtDisplayName = !DisplayName
                pWFTypeForm.TxtWFNo = !WFNo
                pWFTypeForm.Show
            End If
        End With
        
        Set RstWFType = Nothing
    Else
        pWFTypeForm.ClearForm
        pWFTypeForm.Show
    End If
End Sub

' ===============================================================
' Event pWFTypeForm_Delete
' Form event to delete workflow type
'---------------------------------------------------------------
Private Sub pWFTypeForm_Delete(WFNo As String)
    DB.Execute "DELETE FROM TblWorkflowType WHERE WFNo = " & WFNo
End Sub

' ===============================================================
' Event pWFTypeForm_Update
' Form event to update workflow type
'---------------------------------------------------------------
Private Sub pWFTypeForm_Update(WFNo As String)
    Dim RstWFType As Recordset
    
    If WFNo = "" Then
        pWFTypeForm_CreateNew
    Else
        Set RstWFType = ModDatabase.SQLQuery("SELECT * FROM TblWorkflowType WHERE WFNo = " & WFNo)
        
        With RstWFType
            If .RecordCount = 0 Then
                pWFTypeForm_CreateNew
            Else
                .Edit
                !Description = pWFTypeForm.TxtDescription
                !DisplayName = pWFTypeForm.TxtDisplayName
                .Update
            End If
        End With
        
        Set RstWFType = Nothing
    End If
End Sub

' ===============================================================
' Event pWFTypeForm_CreateNew
' Form event to create new workflow type
'---------------------------------------------------------------
Private Sub pWFTypeForm_CreateNew()
    Dim RstWFType As Recordset
    Dim RstMaxWFType As Recordset
    
    Set RstWFType = ModDatabase.SQLQuery("TblWorkflowType")
    Set RstMaxWFType = ModDatabase.SQLQuery("SELECT Max(WFNo) as MaxWFNo FROM TblWorkflowType")
    
    With RstWFType
        .AddNew
        !WFNo = .Fields(0) + 1
        !Description = pWFTypeForm.TxtDescription
        !DisplayName = pWFTypeForm.TxtDisplayName
        !WFName = Replace(!DisplayName, " ", "")
        .Update
    End With
    
    Set RstWFType = Nothing
    Set RstMaxWFType = Nothing
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsProject)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pWorkflows = New Collection
    Set pWFTypeForm = New FrmAdnWFTypes
End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing
    Set pWFTypeForm = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
    Set pWFTypeForm = Nothing
End Sub ' ===============================================================
