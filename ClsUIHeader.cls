VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIHeader
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Feb 17
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pStyle As ClsUIStyle
Private pText As String
Private pShpHeader As Shape
Private pIcon As Shape
Private pVisible As Boolean
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpHeader.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpHeader.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpHeader.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpHeader.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpHeader.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpHeader.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpHeader.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpHeader.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpHeader.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpHeader.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get ShpHeader() As Shape
    Set ShpHeader = pShpHeader
End Property

Public Property Let ShpHeader(ByVal vNewValue As Shape)
    Set pShpHeader = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As ClsUIStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As ClsUIStyle)
    With ShpHeader.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With ShpHeader.Line
        .ForeColor.RGB = vNewValue.BorderColour
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With ShpHeader.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
        If vNewValue.FontVJust <> 0 Then .VerticalAlignment = vNewValue.FontVJust
    End With
    
    With ShpHeader.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
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
    pShpHeader.Visible = pVisible
    If Not pIcon Is Nothing Then pIcon.Visible = pVisible
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pText
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpHeader.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Icon() As Shape
    Set Icon = pIcon
End Property

Public Property Let Icon(ByVal vNewValue As Shape)
    Set pIcon = vNewValue
End Property

' ===============================================================
' Method ReOrder
' Re-orders the shapes so that they are in the right stack order
'---------------------------------------------------------------
Public Sub ReOrder()
    pShpHeader.ZOrder msoBringToFront
    If Not pIcon Is Nothing Then pIcon.ZOrder msoBringToFront
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pShpHeader = ShtMain.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    pShpHeader.Visible = msoFalse
End Sub

' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    pShpHeader.Delete
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pShpHeader = Nothing
End Sub
'---------------------------------------------------------------

