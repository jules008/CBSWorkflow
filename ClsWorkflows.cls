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
Public Sub Add(Workflow As ClsWorkflow)
    Workflow.Initialise Me.Parent
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
                pWorkflows.Add Workflow, Key:=CStr(Workflow.WorkflowNo)
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
Public Sub Initialize(obj As ClsProject)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pWorkflows = New Collection

End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pWorkflows = Nothing
End Sub
' ===============================================================
