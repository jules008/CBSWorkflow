VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUICell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUICell
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 May 20
'===============================================================

Option Explicit
Private pTop As Integer
Private pRow As Integer
Private pCol As Integer
Private pHeight As Integer
Private pLeft As Integer
Private pWidth As Integer
Private pText As String
Private pStyle As ClsUIStyle
Private pName As String
Private pMenu As ClsUIMenu
Private pOnAction As String
Private pBadge As Shape
Private pShpCell As Shape
Private pVisible As Boolean
Private pParent As ClsUITable

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpCell.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpCell.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Row() As Integer
    Row = pRow
End Property

Public Property Let Row(ByVal vNewValue As Integer)
    pRow = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Col() As Integer
    Col = pCol
End Property

Public Property Let Col(ByVal vNewValue As Integer)
    pCol = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpCell.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpCell.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpCell.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpCell.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpCell.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpCell.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pShpCell.TextFrame.Characters.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpCell.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpCell.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpCell.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal vNewValue As ClsUIMenu)
    Set pMenu = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Badge() As Shape
    Set Badge = pBadge
End Property

Public Property Let Badge(ByVal vNewValue As Shape)
    Set pBadge = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpCell.OnAction
End Property

Public Property Let OnAction(ByVal vNewValue As String)
    pOnAction = vNewValue
    pShpCell.OnAction = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpCell.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    
    pVisible = vNewValue
    
    pShpCell.Visible = pVisible
       
End Property

'---------------------------------------------------------------
Public Property Get ShpCell() As Shape
    Set ShpCell = pShpCell
End Property

Public Property Let ShpCell(ByVal vNewValue As Shape)
    Set pShpCell = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As ClsUIStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As ClsUIStyle)
    Set pStyle = vNewValue
    
    With pShpCell.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With pShpCell.Line
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
        .ForeColor.RGB = vNewValue.BorderColour
    End With
    
    With pShpCell.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        If vNewValue.TextDir <> 0 Then .Orientation = vNewValue.TextDir
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
        If vNewValue.FontVJust <> 0 Then .VerticalAlignment = vNewValue.FontVJust
    End With
    
    With pShpCell.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUITable
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUITable)
    Set pParent = obj
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the Cells sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    
'    pShpCell.ZOrder msoBringToFront
    If Not pBadge Is Nothing Then pBadge.ZOrder msoBringToFront
    
    If Not pMenu Is Nothing Then pMenu.ReOrder
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUITable)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pShpCell = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
    Set pStyle = New ClsUIStyle
End Sub

' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    Set pStyle = Nothing
    pShpCell.Delete
    
    If Not pBadge Is Nothing Then
        pBadge.Delete
        Set pBadge = Nothing
    End If
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pShpCell = Nothing
    Set pStyle = Nothing

    If Not pBadge Is Nothing Then
        Set pBadge = Nothing
    End If
End Sub

'---------------------------------------------------------------
