VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIFrames
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Feb 17
'===============================================================

Option Explicit
Private pUIFrames As Collection
Private pParent As ClsUIScreen

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIScreen
    Set Parent = pParent
End Property

Friend Property Set Parent(obj As ClsUIScreen)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIFrames.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUIFrame
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUIFrames.item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UIFrame As ClsUIFrame, Name As String)
    UIFrame.Name = Name
    UIFrame.Initialize Me.Parent
    pUIFrames.Add UIFrame, Key:=CStr(UIFrame.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Frame As ClsUIFrame)
    On Error Resume Next
    pUIFrames.Remove Frame.Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIFrames.Count
End Function

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Frame As ClsUIFrame
        
    For Each Frame In pUIFrames
        Frame.ShpFrame.ZOrder msoBringToFront
        Frame.ReOrder
    Next
    
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIScreen)
    Set Me.Parent = obj
End Sub
Private Sub Class_Initialize()
    Set pUIFrames = New Collection
End Sub
' ===============================================================
Public Sub Terminate()
    Dim Frame As ClsUIFrame

    For Each Frame In pUIFrames
        Frame.Terminate
    Next
    Set pParent = Nothing
    Set pUIFrames = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUIFrames = Nothing
End Sub

'---------------------------------------------------------------
