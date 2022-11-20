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
Private pStyle As ClsUIStyle
Private pDashObjs As ClsUIDashObjs
Private pMenu As ClsUIMenu
Private pEnableHeader As Boolean
Private pHeader As ClsUIHeader
Private pCellHeaders As ClsUICells
Private pButtons As ClsUIButtons
Private pCells As ClsUICells
Private pShpFrame As Shape
Private pVisible As Boolean
Private pZOrder As Integer
Private pTable As ClsUITable
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
Public Property Get CellHeaders() As ClsUICells
    Set CellHeaders = pCellHeaders
End Property

Public Property Let CellHeaders(ByVal vNewValue As ClsUICells)
    Set pCellHeaders = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Cells() As ClsUICells
    Set Cells = pCells
End Property

Public Property Let Cells(ByVal vNewValue As ClsUICells)
    Set pCells = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Buttons() As ClsUIButtons
    Set Buttons = pButtons
End Property

Public Property Let Buttons(ByVal vNewValue As ClsUIButtons)
    Set pButtons = vNewValue
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
Public Property Get ZOrder() As Integer
    ZOrder = pZOrder
End Property

Public Property Let ZOrder(ByVal vNewValue As Integer)
    pZOrder = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Table() As ClsUITable
    Set Table = pTable
End Property

Public Property Let Table(ByVal vNewValue As ClsUITable)
    Set pTable = vNewValue
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
Public Property Get Style() As ClsUIStyle
    Set Style = pStyle
End Property

Public Property Let Style(vNewValue As ClsUIStyle)
    Set pStyle = vNewValue
    
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
        If vNewValue.FontVJust <> 0 Then .VerticalAlignment = vNewValue.FontVJust
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
    Dim UICell As ClsUICell
    
    pShpFrame.ZOrder msoBringToFront
    pHeader.ReOrder
    pTable.ReOrder
    
    For Each DashObj In DashObjs
        DashObj.ShpDashObj.ZOrder msoBringToFront
    Next
    
    pButtons.ReOrder
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
    Set pCells = New ClsUICells
    Set pShpFrame = ShtMain.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    Set pTable = New ClsUITable
    Set pStyle = New ClsUIStyle
    Set pButtons = New ClsUIButtons
    
    pDashObjs.Initialize Me
    pMenu.Initialize Me
    pHeader.Initialize Me
    pTable.Initialize Me
    pButtons.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pDashObjs.Terminate
    pMenu.Terminate
    pHeader.Terminate
    pTable.Terminate
    pCells.Terminate
    pButtons.Terminate
    
    pParent.Frames.RemoveItem Me
    
    Set pParent = Nothing
    pShpFrame.Delete

    Set pStyle = Nothing
    Set pShpFrame = Nothing
    Set pDashObjs = Nothing
    Set pMenu = Nothing
    Set pHeader = Nothing
    Set pTable = Nothing
    Set pButtons = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pStyle = Nothing
    Set pShpFrame = Nothing
    Set pDashObjs = Nothing
    Set pMenu = Nothing
    Set pHeader = Nothing
    Set pCellHeaders = Nothing
    Set pCells = Nothing
    Set pTable = Nothing
    Set pButtons = Nothing
End Sub
