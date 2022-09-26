VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUITable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUITable
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================
'---------------------------------------------------------------
Option Explicit

Private pCells As ClsUICells
Private pTop As Integer
Private pLeft As Integer
Private pNoRows As Integer
Private pNoCols As Integer
Private pRowHeights() As Integer
Private pColWidths() As Integer
Private pText() As String
Private pStyle() As TypeStyle
Private pCellVisible() As Boolean
Private pTableVisible As Boolean
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Cells() As ClsUICells
    Set Cells = pCells
End Property

Public Property Let Cells(ByVal vNewValue As ClsUICells)
    Set pCells = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    If Not Parent Is Nothing Then
        Top = Parent.Top + pTop
    Else
        Top = pTop
    End If
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
    If Not Parent Is Nothing Then
        Left = Parent.Left + pLeft
    Else
        Top = pTop
    End If
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoRows() As Integer
    NoRows = pNoRows
End Property

Public Property Let NoRows(ByVal vNewValue As Integer)
    
    pNoRows = vNewValue
    
    If NoCols > 0 Then
        BuildCells
    End If
    
End Property

'---------------------------------------------------------------
Public Property Get NoCols() As Integer
    NoCols = pNoCols
End Property

Public Property Let NoCols(ByVal vNewValue As Integer)
    
    pNoCols = vNewValue
    
    If NoRows > 0 Then
        BuildCells
    End If
    
End Property

'---------------------------------------------------------------
Public Property Get RowHeights() As Integer()
    RowHeights = pRowHeights
End Property

Public Property Let RowHeights(ByRef vNewValue() As Integer)
    
    pRowHeights = vNewValue
    
    BuildCells
End Property

'---------------------------------------------------------------
Public Property Get ColWidths() As Integer()
    ColWidths = pColWidths
End Property

Public Property Let ColWidths(ByRef vNewValue() As Integer)
    
    pColWidths = vNewValue
    
    BuildCells
End Property

'---------------------------------------------------------------
Public Property Get Text(x As Integer, y As Integer) As String
    Text = pCells(x, y).Text
End Property

Public Property Let Text(x As Integer, y As Integer, vNewValue As String)
    pCells(x, y).Text = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style(x As Integer, y As Integer) As TypeStyle
    Style = pCells(x, y).Style
End Property

Public Property Let Style(x As Integer, y As Integer, vNewValue As TypeStyle)
    pStyle(x, y) = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CellVisible(x As Integer, y As Integer) As Boolean
    CellVisible = pCells(x, y).Visible
End Property

Public Property Let CellVisible(x As Integer, y As Integer, vNewValue As Boolean)
    pCells(x, y).Visible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TableVisible() As Boolean
    TableVisible = pTableVisible
End Property

Public Property Let TableVisible(ByVal vNewValue As Boolean)
    Sin 0
    For Each Cell In pCells
        Cell.Visible = vNewValue
    Next
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property
'---------------------------------------------------------------

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Cell As ClsUICell
    
    For Each Cell In pCells
        With Cell
            .ShpCell.ZOrder msoBringToFront
            If Not .Badge Is Nothing Then .Badge.ZOrder msoBringToFront
        End With
    Next
    
'    For Each Cell In CellHeaders
'        With Cell
'            .ShpCell.ZOrder msoBringToFront
'            If Not .Badge Is Nothing Then .Badge.ZOrder msoBringToFront
'        End With
'    Next
End Sub

' ===============================================================
' Method BuildCells
' Builds matrix of cells once dimensions are received
'---------------------------------------------------------------
Private Sub BuildCells()
    Dim x, y As Integer
    Dim Cell As ClsUICell
    
    For x = 1 To pNoCols
        For y = 1 To pNoRows
            Set Cell = New ClsUICell
            
            With Cell
                .Left = .Parent.Left
                .Top = .Parent.Top
                .Col = x
                .Row = y
            End With
            
            pCells.AddItem Cell
        Next
    Next
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pCells = New ClsUICells
    pCells.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing
    pCells.Terminate
    
    Set pCells = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pCells = Nothing
End Sub
' ===============================================================


