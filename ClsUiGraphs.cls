VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUiGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIGraphs
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Dec 20
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
'===============================================================

Option Explicit
Private pUIGraphs As Collection
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIGraphs.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------
Public Function FindItem(Name As Variant) As ClsUIGraph
    On Error Resume Next
    Set FindItem = pUIGraphs.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(UIGraph As ClsUIGraph)
    UIGraph.Initialize Me.Parent
    pUIGraphs.Add UIGraph, Key:=CStr(UIGraph.Name)
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUIGraphs.Remove Name
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Graph As ClsUIGraph
        
    For Each Graph In pUIGraphs
        Graph.ReOrder
    Next
    
End Sub

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIGraphs.Count
End Function

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub
Private Sub Class_Initialize()
    Set pUIGraphs = New Collection

End Sub
' ===============================================================
Public Sub Terminate()
    Dim Graph As ClsUIGraph

    For Each Graph In pUIGraphs
        Graph.Terminate
    Next
        
    Set pParent = Nothing
    Set pUIGraphs = Nothing

End Sub
Private Sub Class_Terminate()
    Set pUIGraphs = Nothing
End Sub
' ===============================================================

