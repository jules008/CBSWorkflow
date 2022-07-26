VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIScreen
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 16 Apr 20
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pStyle As ClsUIStyle
Private pShpScreen As Shape
Private pFrames As ClsUIFrames
Private pVisible As Boolean

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpScreen.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpScreen.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpScreen.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpScreen.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpScreen.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpScreen.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpScreen.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpScreen.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpScreen.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpScreen.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get ShpScreen() As Shape
    Set ShpScreen = pShpScreen
End Property

Public Property Let ShpScreen(ByVal vNewValue As Shape)
    Set pShpScreen = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Frames() As ClsUIFrames
    Set Frames = pFrames
End Property

Public Property Let Frames(ByRef vNewValue As ClsUIFrames)
    Set pFrames = vNewValue
End Property
'
''---------------------------------------------------------------
'Public Property Get Menu() As ClsUIMenu
'    Set Menu = pMenu
'End Property
'
'Public Property Let Menu(ByVal vNewValue As ClsUIMenu)
'    Set pMenu = vNewValue
'End Property

'---------------------------------------------------------------
Public Property Get Style() As ClsUIStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As ClsUIStyle)
    Set pStyle = vNewValue
    
    With pShpScreen.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = pStyle.Fill1
        .BackColor.RGB = pStyle.Fill2
    End With
    
    With pShpScreen.Line
        .ForeColor.RGB = pStyle.BorderColour
        .Weight = pStyle.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpScreen.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = pStyle.FontStyle
        .Characters.Font.Color = pStyle.FontColour
        .Characters.Font.Size = pStyle.FontSize
        If pStyle.FontXJust <> 0 Then .HorizontalAlignment = pStyle.FontXJust
        If pStyle.FontVJust <> 0 Then .VerticalAlignment = pStyle.FontVJust
    End With
    
    With pShpScreen.Shadow
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
    Visible = pShpScreen.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    pShpScreen.Visible = pVisible
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Frame As ClsUIFrame
    Dim i As Integer
    
    pShpScreen.ZOrder msoSendToBack
    
    For i = 3 To 0 Step -1
    For Each Frame In pFrames
            If i = Frame.ZOrder Then
                'Debug.Print Frame.Name, Frame.ZOrder
        Frame.ReOrder
            End If
        Next
    Next
        
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pFrames = New ClsUIFrames
    Set pStyle = New ClsUIStyle
    Set pShpScreen = ShtMain.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    
    pFrames.Initialize Me
    
End Sub
' ===============================================================
Public Sub Terminate()
    pFrames.Terminate
    pShpScreen.Delete

    Set pFrames = Nothing
    Set pShpScreen = Nothing
    Set pStyle = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pFrames = Nothing
    Set pShpScreen = Nothing
    Set pStyle = Nothing
    
End Sub

'---------------------------------------------------------------


