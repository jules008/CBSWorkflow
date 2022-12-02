VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsProjects
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Sep 22
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

Private pProjects As Collection
Private pParent As ClsSPV

'---------------------------------------------------------------
Public Property Get Parent() As ClsSPV
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsSPV)
    Set pParent = obj
End Property
'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pProjects.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(ProjectNo As Variant) As ClsProject
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pProjects.Item(ProjectNo)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(ByRef Project As ClsProject)
    Project.Initialize Me.Parent
    pProjects.Add Project, Key:=CStr(Project.ProjectNo)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(ProjectNo As Variant)
    pProjects(ProjectNo).Parent = Nothing
    pProjects.Remove ProjectNo
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pProjects.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
' ---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstProject As Recordset
    Dim Project As ClsProject
    Dim i As Integer

    Set RstProject = ModDatabase.SQLQuery("TblProject")
    With RstProject
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Project = New ClsProject
                Project.DBGet !ProjectNo
                pProjects.Add Project, Key:=CStr(Project.ProjectNo)
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
    Dim Project As ClsProject
    For Each Project In pProjects
        Project.DBSave
    Next
End Sub

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim Project As ClsProject
    For Each Project In pProjects
        pProjects.Remove CStr(Project.ProjectNo)
    Next
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Project As ClsProject
    For Each Project In pProjects
        pProjects.Remove CStr(Project.ProjectNo)
        Project.DBDelete
    Next
End Sub

' ===============================================================
' Method Destroy
' Removes item from collection and destroys it
' ---------------------------------------------------------------
Public Sub Destroy(ProjectNo As Variant)
    Dim Project As ClsProject

    Set Project = pProjects(ProjectNo)
    Project.Parent = Nothing
    Project.Terminate
    pProjects.Remove ProjectNo
    Project.DBDelete True

    Set Project = Nothing
End Sub
' ===============================================================
Public Sub Initialize(obj As ClsSPV)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pProjects = New Collection

End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pProjects = Nothing
End Sub
' ===============================================================
