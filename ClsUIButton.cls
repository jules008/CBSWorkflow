VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIButton
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 15 Jun 20
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pText As String
Private pIcon As Shape
Private pIconLeft As Integer
Private pIconTop As Integer
Private pBadge As ClsUICell
Private pBadgeLeft As Integer
Private pBadgeTop As Integer
Private pBadgeText As String
Private pButtonLvl As Integer
Private pButtonIndex As String
Private pSelectStyle As ClsUIStyle
Private pUnSelectStyle As ClsUIStyle
Private pShpButton As Shape
Private pSelected As Boolean
Private pVisible As Boolean
Private pOnAction As String
Private pParent As ClsUIMenu

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpButton.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpButton.Top = pTop
    If Not Icon Is Nothing Then
        Icon.Top = pIconTop + pTop
    End If
    If Not Badge Is Nothing Then
        Badge.Top = pBadgeTop + pTop
    End If
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpButton.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpButton.Left = pLeft
    If Not Icon Is Nothing Then
        Icon.Left = pIconLeft + pLeft
    End If
    If Not Badge Is Nothing Then
        Badge.Left = pBadgeLeft + pLeft
    End If
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpButton.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpButton.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpButton.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpButton.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpButton.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpButton.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pShpButton.TextFrame.Characters.Caption
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpButton.TextFrame.Characters.Caption = pText
End Property

'---------------------------------------------------------------
Public Property Get Icon() As Shape
    Set Icon = pIcon
End Property

Public Property Let Icon(ByVal vNewValue As Shape)
    Set pIcon = vNewValue
End Property

Public Property Let IconLeft(ByVal IconLeft As Integer)
    pIconLeft = IconLeft
    pIcon.Left = pLeft + IconLeft
End Property

Public Property Let IconTop(ByVal IconTop As Integer)
    pIconTop = IconTop
    pIcon.Top = pTop + IconTop
End Property

Public Property Let BadgeLeft(ByVal BadgeLeft As Integer)
    pBadgeLeft = BadgeLeft
    pBadge.Left = pLeft + BadgeLeft
End Property

Public Property Let BadgeTop(ByVal BadgeTop As Integer)
    pBadgeTop = BadgeTop
    pBadge.Top = pTop + BadgeTop
End Property

'---------------------------------------------------------------
Public Property Get Badge() As ClsUICell
    Set Badge = pBadge
End Property

Public Property Let Badge(ByVal vNewValue As ClsUICell)
    Set pBadge = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get BadgeText() As String
    BadgeText = pBadgeText
End Property

Public Property Let BadgeText(ByVal vNewValue As String)
    pBadgeText = vNewValue
    
    If pBadgeText <> "0" Then
        pBadge.Visible = msoCTrue
    Else
        pBadge.Visible = msoFalse
    End If
    
    If CInt(pBadgeText) > 99 Then pBadgeText = "99"
    
    pBadge.Text = pBadgeText
End Property

'---------------------------------------------------------------
Public Property Get ButtonLvl() As Integer
    Dim AryIndex
    
    AryIndex = Split(ButtonIndex, ".")
    pButtonLvl = UBound(AryIndex) + 1
    ButtonLvl = pButtonLvl
    
End Property

'---------------------------------------------------------------
Public Property Get ButtonIndex() As String
    ButtonIndex = pButtonIndex
End Property

Public Property Let ButtonIndex(ByVal vNewValue As String)
    pButtonIndex = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SelectStyle() As ClsUIStyle
    Set SelectStyle = pSelectStyle
End Property

Public Property Let SelectStyle(vNewValue As ClsUIStyle)
    Set pSelectStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UnSelectStyle() As ClsUIStyle
    Set UnSelectStyle = pUnSelectStyle
End Property

Public Property Let UnSelectStyle(vNewValue As ClsUIStyle)
    Set pUnSelectStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ShpButton() As Shape
    Set ShpButton = pShpButton
End Property

Public Property Let ShpButton(ByVal vNewValue As Shape)
    Set pShpButton = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Selected() As Boolean
    Selected = pSelected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    Dim Style As ClsUIStyle
    
    pSelected = vNewValue
        
    If pSelected Then
        Set Style = pSelectStyle
    Else
        Set Style = pUnSelectStyle
    End If
    
    With pShpButton.Fill
'        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = Style.Fill1
        .BackColor.RGB = Style.Fill2
    End With
    
    With pShpButton.Line
        .Weight = Style.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
        .ForeColor.RGB = Style.BorderColour
    End With
    
    With pShpButton.TextFrame
        .Characters.Font.Bold = Style.FontBold
        .Characters.Font.Name = Style.FontStyle
        .Characters.Font.Color = Style.FontColour
        .Characters.Font.Size = Style.FontSize
        If Style.FontXJust <> 0 Then .HorizontalAlignment = Style.FontXJust
'        If Style.FontVJust <> 0 Then .VerticalAlignment = Style.FontVJust
    End With
    
    With pShpButton.Shadow
        If Style.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = Style.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    pShpButton.Visible = vNewValue
    If Not pIcon Is Nothing Then pIcon.Visible = vNewValue
    If Not pBadge Is Nothing Then pBadge.Visible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpButton.OnAction
End Property

Public Property Let OnAction(ByVal vNewValue As String)
    pShpButton.OnAction = vNewValue
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    pShpButton.ZOrder msoBringToFront
    If Not pIcon Is Nothing Then pIcon.ZOrder msoBringToFront
    If Not pBadge Is Nothing Then pBadge.ReOrder
End Sub

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIMenu
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIMenu)
    Set pParent = obj
End Property

' ===============================================================
Public Sub Initialize(obj As ClsUIMenu)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pShpButton = ShtMain.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
End Sub
' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    pShpButton.Delete
    Set pShpButton = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pShpButton = Nothing
End Sub


