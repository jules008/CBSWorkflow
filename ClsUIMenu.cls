VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIMenu
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 15 Apr 20
'===============================================================

Option Explicit
Private pUIMenu As Collection
Private pTop As Integer
Private pLeft As Integer
Private pVisible As Integer
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    Dim Button As ClsUIButton
    
    pVisible = vNewValue
    
    For Each Button In pUIMenu
        Button.Visible = vNewValue
    Next
    
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIMenu.[_NewEnum]
End Function

' ===============================================================
' Method FindButton
' Finds Button from index
'---------------------------------------------------------------
Public Function FindButton(Name As Variant) As ClsUIButton
Attribute FindButton.VB_UserMemId = 0
    On Error Resume Next
    Set FindButton = pUIMenu.Item(Name)
End Function

' ===============================================================
' Method AddButton
' Add Button to Collection
'---------------------------------------------------------------
Public Sub AddButton(UIButton As ClsUIButton)
    pUIMenu.Add UIButton, Key:=CStr(UIButton.Name)
End Sub

' ===============================================================
' Method RemoveButton
' Deletes Button from collection
'---------------------------------------------------------------
Public Sub RemoveButton(Name As Variant)
    pUIMenu.Remove Name
End Sub

' ===============================================================
' Method ButtonCount
' Returns number of Buttons in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIMenu.Count
End Function

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Button As ClsUIButton
        
    For Each Button In pUIMenu
        Button.ShpButton.ZOrder msoBringToFront
        Button.ReOrder
    Next
    
End Sub

' ===============================================================
' Method ButtonClick
' Presses selected button of menu
'---------------------------------------------------------------
Public Sub ButtonClick(BtnName)
    Dim Button As ClsUIButton
        
    For Each Button In pUIMenu
        If Button.Name = pUIMenu(BtnName).Name Then
            Button.Selected = True
        Else
            Button.Selected = False
        End If
    Next
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pUIMenu = New Collection
End Sub

' ===============================================================
Public Sub Terminate()
    Dim Button As ClsUIButton
        
    For Each Button In pUIMenu
        Button.Terminate
    Next
    Set pParent = Nothing
    Set pUIMenu = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUIMenu = Nothing
End Sub

'---------------------------------------------------------------
