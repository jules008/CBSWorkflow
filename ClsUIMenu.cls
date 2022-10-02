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
Public Function FindButton(ButtonIndex As String) As ClsUIButton
Attribute FindButton.VB_UserMemId = 0
    On Error Resume Next
    Set FindButton = pUIMenu.Item(ButtonIndex)
End Function

' ===============================================================
' Method AddButton
' Add Button to Collection
'---------------------------------------------------------------
Public Sub AddButton(UIButton As ClsUIButton)
    Dim i As Integer
    Dim AryIndex() As String
    Dim ParentIndex As Integer
    Dim HorOffset As Integer
    Dim ChildNo As Integer
    
    AryIndex = Split(UIButton.ButtonIndex, ".")
    
    ParentIndex = AryIndex(0)
    ChildNo = AryIndex(UBound(AryIndex))
        
'    Debug.Print "Button Index: "; UIButton.ButtonIndex
'    Debug.Print "Parent Index: "; ParentIndex
'    Debug.Print "Child No: "; ChildNo
    
    'treat the first button differently as the collection is empty
    If pUIMenu.Count = 0 Then
        With UIButton
            .Top = pTop
            .Left = pLeft
'            Debug.Print "i: "; i
'            Debug.Print "Button Top: "; UIButton.Top
'            Debug.Print "Button Left: "; UIButton.Left
        End With
    End If
        
    With UIButton
        .OnAction = "'ModUIMenu.ButtonClickEvent(" & .ButtonIndex & ")'"
        If .ButtonLvl = 1 Then
            .Visible = True
            .Top = pTop + (.ButtonIndex - 1) * .Height
            .Left = pLeft
        Else
            .Visible = False
            
            'get y pos from parent button
            .Top = FindButton(CStr(ParentIndex)).Top + (CInt(ChildNo) - 1) * .Height
            .Left = pLeft + (.ButtonLvl - 1) * .Width
        End If
        .Selected = False
'        Debug.Print "i: "; i
'        Debug.Print "Button Top: "; UIButton.Top
'        Debug.Print "Button Left: "; UIButton.Left
    End With
        
    Debug.Print
    
    pUIMenu.Add UIButton, Key:=CStr(UIButton.ButtonIndex)
End Sub

' ===============================================================
' Method RemoveButton
' Deletes Button from collection
'---------------------------------------------------------------
Public Sub RemoveButton(ButtonIndex As String)
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
Public Sub ButtonClick(BtnName As String)
    Dim Button As ClsUIButton
    Dim SubButton As ClsUIButton
    Dim SubMenu As Boolean
        
    For Each Button In pUIMenu
        With Button
            If .Name = pUIMenu(BtnName).Name Then
                .Selected = True
                
                Debug.Print .OnAction
                'if it isn't a main menu button, show sub menu group
                Dim i As Integer
                i = 1
                Do
                    Set SubButton = FindButton(.ButtonIndex & "." & i)
                    If Not SubButton Is Nothing Then
                        SubButton.Visible = True
                         SubMenu = True
                    Else
                        Exit Do
                    End If
                    If i = 10 Then Exit Do
                    i = i + 1
                Loop
                
                If Not SubMenu Then
                    
                    'main button has been clicked, so hide all sub menus and process click
                    Application.Wait 1000
                    For Each SubButton In pUIMenu
                        If SubButton.ButtonLvl > 1 Then SubButton.Visible = False
                    Next
                    ModUIButtons.ProcessBtnClicks Replace(.ButtonIndex, ".", "")
                
                End If
        Else
                .Selected = False
        End If
        End With
    Next
    Set SubButton = Nothing
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
