VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIDashObjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIDashObjs
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 21 May 20
'===============================================================

Option Explicit
Private pUIDashObjs As Collection
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
    Set NewEnum = pUIDashObjs.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUIDashObj
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUIDashObjs.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(UIDashObj As ClsUIDashObj)
    If UIDashObj Is Nothing Then
        Set UIDashObj = New ClsUIDashObj
    End If
    UIDashObj.Initialize Me.Parent
    pUIDashObjs.Add UIDashObj, Key:=CStr(UIDashObj.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUIDashObjs.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIDashObjs.Count
End Function

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set pUIDashObjs = New Collection
    Set Me.Parent = obj

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

'---------------------------------------------------------------
