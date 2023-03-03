VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIBadge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIBadge
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 26 Feb 23
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pText As String
Private pLocked As Boolean
Private pShpBadge As Shape
Private pVisible As Boolean
Private pRotation As Integer
Private pOnAction As String
Private pZOrder As MsoZOrderCmd
Private pStyle As ClsUIStyle
Private pParent As ClsUICell

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpBadge.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpBadge.Top = pTop + pParent.Top
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpBadge.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpBadge.Left = pLeft + pParent.Left
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpBadge.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpBadge.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpBadge.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpBadge.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpBadge.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    pText = pShpBadge.TextFrame.Characters.Text
    Text = pText
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpBadge.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Locked() As Boolean
    pLocked = pShpBadge.ControlFormat.LockedText = Locked
    Locked = pLocked
End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
    pLocked = vNewValue
    pShpBadge.Locked = True
    pShpBadge.ControlFormat.LockedText = Locked
End Property

'---------------------------------------------------------------
Public Property Get ShpBadge() As Shape
    Set ShpBadge = pShpBadge
End Property

Public Property Let ShpBadge(ByVal vNewValue As Shape)
    Set pShpBadge = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpBadge.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    If Not pShpBadge Is Nothing Then pShpBadge.Visible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Rotation() As Integer
    Rotation = pShpBadge.Rotation
End Property

Public Property Let Rotation(ByVal vNewValue As Integer)
    pShpBadge.Rotation = vNewValue
    pRotation = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpBadge.OnAction
End Property

Public Property Let OnAction(ByVal vNewValue As String)
    pShpBadge.OnAction = vNewValue
    pOnAction = vNewValue
End Property

Public Property Let ZOrder(ByVal vNewValue As MsoZOrderCmd)
    pShpBadge.ZOrder vNewValue
    pZOrder = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As ClsUIStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As ClsUIStyle)
    Set pStyle = vNewValue
    
    With pShpBadge.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With pShpBadge.Line
        .ForeColor.RGB = vNewValue.BorderColour
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpBadge.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        .MarginTop = vNewValue.MarginTop
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
        If vNewValue.FontVJust <> 0 Then .VerticalAlignment = vNewValue.FontVJust
    End With
    
    With pShpBadge.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property
'---------------------------------------------------------------
Public Property Get Parent() As ClsUICell
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUICell)
    Set pParent = obj
End Property
'---------------------------------------------------------------
' ===============================================================
Public Sub Initialize(obj As ClsUICell)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

' ===============================================================
Public Sub Terminate()
    pShpBadge.Delete
    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub
' ===============================================================
