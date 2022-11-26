VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIButtons
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 02 Oct 22
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

Private pUIButtons As Collection
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property
'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIButtons.[_NewEnum]
End Function

' ===============================================================
' Method Find
' Finds Item from index
' ---------------------------------------------------------------
Public Function Find(ButtonIndex As Variant) As ClsUIButton
Attribute Find.VB_UserMemId = 0
    On Error Resume Next
    Set Find = pUIButtons.item(ButtonIndex)
End Function

' ===============================================================
' Method Add
' Add item to Collection
' ---------------------------------------------------------------
Public Sub Add(UIButton As ClsUIButton)
    UIButton.Left = pParent.Left + UIButton.Left
    UIButton.Top = pParent.Top + UIButton.Top
    pUIButtons.Add UIButton, Key:=CStr(UIButton.Name)
End Sub

' ===============================================================
' Method Remove
' Removes item from collection
' ---------------------------------------------------------------
Public Sub Remove(ButtonIndex As Variant)
    pUIButtons(ButtonIndex).Parent = Nothing
    pUIButtons.Remove ButtonIndex
End Sub

' ===============================================================
' Method Count
' Returns number of items in collection
' ---------------------------------------------------------------
Public Function Count() As Integer
    Count = pUIButtons.Count
End Function

' ===============================================================
' Method RemoveCollection
' Deletes all items in collection
' ---------------------------------------------------------------
Public Sub RemoveCollection()
    Dim UIButton As ClsUIButton
    For Each UIButton In pUIButtons
        pUIButtons.Remove CStr(UIButton.ButtonIndex)
    Next
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Button As ClsUIButton
    
    For Each Button In pUIButtons
        Button.ReOrder
    Next
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pUIButtons = New Collection

End Sub
' ===============================================================
Public Sub Terminate()
    Dim Button As ClsUIButton
    
    For Each Button In pUIButtons
        Button.ShpButton.Delete
        If Not Button.Badge Is Nothing Then Button.Badge.ShpCell.Delete
        If Not Button.Icon Is Nothing Then Button.Icon.Delete
    Next
    
    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUIButtons = Nothing
End Sub
' ===============================================================
