VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIBadges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIBadges
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Feb 23
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

Private pUIBadges As Collection
Private pParent As ClsUICell

'---------------------------------------------------------------
Public Property Get Parent() As ClsUICell
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUICell)
    Set pParent = obj
End Property

'---------------------------------------------------------------
Public Property Let SetLeft(BadgeName As String, ByVal Left As Integer)
    pUIBadges(BadgeName).Left = Parent.Left + Left
End Property

'---------------------------------------------------------------
Public Property Let SetTop(BadgeName As String, ByVal Top As Integer)
    pUIBadges(BadgeName).Top = Parent.Top + Top
End Property

'---------------------------------------------------------------
Public Property Let SetWidth(BadgeName As String, ByVal Width As Integer)
    pUIBadges(BadgeName).Width = Width
End Property

'---------------------------------------------------------------
Public Property Let SetHeight(BadgeName As String, ByVal Height As Integer)
    pUIBadges(BadgeName).Height = Height
End Property

' ---------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIBadges.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(Name As Variant) As ClsUIBadge
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pUIBadges.Item(Name)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(UIBadge As ClsUIBadge)
    pUIBadges.Add UIBadge, Key:=CStr(UIBadge.Name)
    UIBadge.Initialize Me.Parent
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(Name As ClsUIBadge)
    pUIBadges(Name).Parent = Nothing
    pUIBadges.Remove Name.Name
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pUIBadges.Count
End Function


' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim UIBadge As ClsUIBadge
    For Each UIBadge In pUIBadges
        pUIBadges.Remove CStr(UIBadge.Name)
    Next
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the Cells sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Badge As ClsUIBadge
    
    For Each Badge In pUIBadges
        Badge.ZOrder = msoBringToFront
    Next
    
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUICell)
    Set Me.Parent = obj
End Sub

'---------------------------------------------------------------
Private Sub Class_Initialize()
    Set pUIBadges = New Collection

End Sub

' ===============================================================
Public Sub Terminate()
    Dim Badge As ClsUIBadge
    
    For Each Badge In pUIBadges
        Badge.Terminate
        Set Badge = Nothing
    Next
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUIBadges = Nothing
End Sub
' ===============================================================
