VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIFrame
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
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pStyle As TypeStyle
'Private pGraphs As ClsUiGraphs
Private pDashObjs As ClsUIDashObjs
Private pMenu As ClsUIMenu
Private pEnableHeader As Boolean
Private pHeader As ClsUIHeader
Private pLineitemHeaders As ClsUILineitems
Private pLineitems As ClsUILineitems
Private pShpFrame As Shape
'Private pMatrixs As ClsUIMatrixs
Private pVisible As Boolean
Private pParent As ClsUIScreen

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)

    pTop = vNewValue + Parent.Top
    pShpFrame.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpFrame.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue + Parent.Left
    pShpFrame.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpFrame.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpFrame.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpFrame.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpFrame.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpFrame.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpFrame.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get DashObjs() As ClsUIDashObjs
    Set DashObjs = pDashObjs
End Property

Public Property Let DashObjs(ByVal vNewValue As ClsUIDashObjs)
    Set pDashObjs = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal vNewValue As ClsUIMenu)
    Set pMenu = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EnableHeader() As Boolean
    EnableHeader = pEnableHeader
End Property

Public Property Let EnableHeader(ByVal vNewValue As Boolean)
    pEnableHeader = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Header() As ClsUIHeader
    Set Header = pHeader
End Property

Public Property Let Header(ByVal vNewValue As ClsUIHeader)
    Set pHeader = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LineitemHeaders() As ClsUILineitems
    Set LineitemHeaders = pLineitemHeaders
End Property

Public Property Let LineitemHeaders(ByVal vNewValue As ClsUILineitems)
    Set pLineitemHeaders = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Lineitems() As ClsUILineitems
    Set Lineitems = pLineitems
End Property

Public Property Let Lineitems(ByVal vNewValue As ClsUILineitems)
    Set pLineitems = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpFrame.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    Dim DashObj As ClsUIDashObj
    
    pVisible = vNewValue
    
    pShpFrame.Visible = pVisible
    
    For Each DashObj In pDashObjs
        DashObj.Visible = pVisible
    Next
    
    If pEnableHeader Then pHeader.Visible = pVisible
    
    pMenu.Visible = pVisible
    
End Property

'---------------------------------------------------------------
Public Property Get ShpFrame() As Shape
    Set ShpFrame = pShpFrame
End Property

Public Property Let ShpFrame(ByVal vNewValue As Shape)
    Set pShpFrame = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIScreen
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIScreen)
    Set pParent = obj
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
    
    With pShpFrame.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With pShpFrame.Line
        .ForeColor.RGB = vNewValue.BorderColour
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpFrame.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
        If vNewValue.FontYJust <> 0 Then .VerticalAlignment = vNewValue.FontYJust
    End With
    
    With pShpFrame.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim DashObj As ClsUIDashObj
    Dim UILineItem As ClsUILineitem
    
    pShpFrame.ZOrder msoBringToFront
    pHeader.ReOrder
    
    For Each DashObj In DashObjs
        DashObj.ShpDashObj.ZOrder msoBringToFront
    Next
    
    For Each UILineItem In Lineitems
        With UILineItem
            .ShpLineitem.ZOrder msoBringToFront
            If Not .Badge Is Nothing Then .Badge.ZOrder msoBringToFront
        End With
    Next
    
    For Each UILineItem In LineitemHeaders
        With UILineItem
            .ShpLineitem.ZOrder msoBringToFront
            If Not .Badge Is Nothing Then .Badge.ZOrder msoBringToFront
        End With
    Next
    
    pMenu.ReOrder
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIScreen)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pDashObjs = New ClsUIDashObjs
    Set pMenu = New ClsUIMenu
    Set pHeader = New ClsUIHeader
    Set pLineitemHeaders = New ClsUILineitems
    Set pLineitems = New ClsUILineitems
    Set pShpFrame = ShtMain.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    
    pLineitems.Initialize Me
    pDashObjs.Initialize Me
    pMenu.Initialize Me
    pHeader.Initialize Me
    pLineitemHeaders.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pDashObjs.Terminate
    pMenu.Terminate
    pHeader.Terminate
    pLineitemHeaders.Terminate
    pLineitems.Terminate
    Set pParent = Nothing
    pShpFrame.Delete

    Set pShpFrame = Nothing
    Set pDashObjs = Nothing
    Set pMenu = Nothing
    Set pHeader = Nothing
    Set pLineitemHeaders = Nothing
    Set pLineitems = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pShpFrame = Nothing
    Set pDashObjs = Nothing
    Set pMenu = Nothing
    Set pHeader = Nothing
    Set pLineitemHeaders = Nothing
    Set pLineitems = Nothing
End Sub

'---------------------------------------------------------------
