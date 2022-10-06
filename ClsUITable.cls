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
Private pHeight As Integer
Private pNoRows As Integer
Private pNoCols As Integer
Private pHPad As Integer
Private pVPad As Integer
Private pSubTableVOff As Integer
Private pSubTableHOff As Integer
Private pRowHeight As Integer
Private pColWidths() As Integer
Private pCellVisible() As Boolean
Private pRstText As Recordset
Private pOnAction() As String
Private pStyles() As String
Private pStylesColl As ClsUIStyles
Private pHeadingText() As String
Private pHeadingStyle As ClsUIStyle
Private pHeadingHeight As Integer
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
        Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    If Not Parent Is Nothing Then
        pTop = Parent.Top + vNewValue
    Else
    pTop = vNewValue
    End If
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    
    If Not Parent Is Nothing Then
        pLeft = Parent.Left + vNewValue
    Else
        pLeft = vNewValue
    End If
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Dim i As Integer
    Dim DefHeight As Integer
    
    pHeight = (NoRows + 1) * (pHeight + pVPad)
    
    Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get NoRows() As Integer
    NoRows = pNoRows
End Property

Public Property Let NoRows(ByVal vNewValue As Integer)
    
    pNoRows = vNewValue
    
    If pNoCols > 0 And pNoRows > 0 Then
        ReDimArrays
    End If
    
End Property

'---------------------------------------------------------------
Public Property Get NoCols() As Integer
    NoCols = pNoCols
End Property

'---------------------------------------------------------------
Public Property Get HPad() As Integer
    HPad = pHPad
End Property
    
Public Property Let HPad(ByVal vNewValue As Integer)
    pHPad = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get VPad() As Integer
    VPad = pVPad
End Property

Public Property Let VPad(ByVal vNewValue As Integer)
    pVPad = vNewValue
End Property
    
Public Property Let SubTableVOff(ByVal vNewValue As Integer)
    pSubTableVOff = vNewValue
End Property
    
Public Property Let SubTableHOff(ByVal vNewValue As Integer)
    pSubTableHOff = vNewValue
End Property

Public Property Let RowHeight(ByRef vNewValue As String)
    pRowHeight = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let ColWidths(ByRef vNewValue As String)
    Dim AryCols() As String
    Dim i As Integer
    
    AryCols = Split(vNewValue, ":")
    
    pNoCols = UBound(AryCols) + 1
    
    ReDim pColWidths(0 To pNoCols - 1)
    
    For i = 0 To UBound(AryCols)
        pColWidths(i) = CInt(AryCols(i))
    Next
    
End Property

'---------------------------------------------------------------
Public Property Get CellVisible() As Boolean()
    CellVisible = pCellVisible
End Property

Public Property Let CellVisible(vNewValue() As Boolean)
    pCellVisible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StylesColl() As ClsUIStyles
    Set StylesColl = pStylesColl
End Property

Public Property Let StylesColl(vNewValue As ClsUIStyles)
    Set pStylesColl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let HeadingText(ByRef vNewValue As String)
    Dim AryText() As String
    Dim i As Integer
        
    AryText = Split(vNewValue, ":")
    
    ReDim pHeadingText(0 To UBound(AryText))
    For i = 0 To UBound(AryText)
        pHeadingText(i) = (AryText(i))
    Next
    
End Property

'---------------------------------------------------------------
Public Property Let HeadingStyle(ByRef vNewValue As ClsUIStyle)
    Set pHeadingStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get HeadingHeight() As Integer
    HeadingHeight = pHeadingHeight
End Property

Public Property Let HeadingHeight(ByVal vNewValue As Integer)
    pHeadingHeight = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TableVisible() As Boolean
    TableVisible = pTableVisible
End Property

Public Property Let TableVisible(ByVal vNewValue As Boolean)
    Dim Cell As ClsUICell
    
    Sin 0
    For Each Cell In pCells
        Cell.Visible = vNewValue
    Next
End Property

'---------------------------------------------------------------
Public Property Get RstText() As Recordset
    Set RstText = pRstText
End Property

Public Property Let RstText(ByVal vNewValue As Recordset)
    Set pRstText = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String()
    OnAction = pOnAction
End Property

Public Property Let OnAction(ByRef vNewValue() As String)
    pOnAction = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Styles() As String()
    Styles = pStyles
End Property

Public Property Let Styles(ByRef vNewValue() As String)
    pStyles = vNewValue
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
' Method ReDimArrays
' Sets dimensions of arrays once no of rows and cols are known
'---------------------------------------------------------------
Private Sub ReDimArrays()
    ReDim pOnAction(0 To pNoCols - 1, 0 To pNoRows)
    ReDim pStyles(0 To pNoCols - 1, 0 To pNoRows)
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Cell As ClsUICell
    
    Sin 0
    For Each Cell In pCells
        With Cell
            .ShpCell.ZOrder msoBringToFront
            If Not .Badge Is Nothing Then .Badge.ZOrder msoBringToFront
        End With
    Next
End Sub

' ===============================================================
' Method BuildCells
' Builds matrix of cells once dimensions are received.
'---------------------------------------------------------------
Public Sub BuildCells(Optional SplitRow As Integer, Optional SplitSize As Integer, Optional SubTable As ClsUITable)
    Dim x, y As Integer
    Dim Cell As ClsUICell
    Dim ColOffset As Integer
    Dim RowOffset As Integer
    
    Debug.Print "Buildcells"
       
    Debug.Assert pNoCols > 0 And pNoRows > 0
    
    If pNoCols > 0 And pNoRows > 0 Then
        pCells.DeleteCollection
        
        'heading
        For x = 0 To pNoCols - 1
            Set Cell = New ClsUICell
            Cell.Name = "Cell - " & x & "-" & "0"
            pCells.AddCell Cell, x, y
            With Cell
                .Left = .Parent.Left + ColOffset
                .Top = .Parent.Top + RowOffset
                .Width = pColWidths(x)
                .Text = pHeadingText(x)
                .OnAction = pOnAction(x, 0)
                .Style = pHeadingStyle
                .Height = pHeadingHeight
            End With
            ColOffset = ColOffset + Cell.Width + HPad
        Next
        ColOffset = 0
        RowOffset = pHeadingHeight
        
        pRstText.MoveFirst
        For y = 1 To pNoRows
            
            If SplitRow <> 0 Then
                If y = SplitRow + 1 Then
                    
'                    SubTable.Initialize Me.Parent
                    
                    With SubTable
                        
                        .Top = pTop + RowOffset + pSubTableHOff
                        .Left = Me.Left + pSubTableVOff
                    
                        .BuildCells
                        RowOffset = RowOffset + .Height + pVPad + (pSubTableHOff * 2)
                    
                    End With
                    
                End If
            End If
            
            For x = 0 To pNoCols - 1
                Set Cell = New ClsUICell
                Cell.Name = "Cell - " & x & "-" & y
                pCells.AddCell Cell, x, y
                With Cell
                    .Left = .Parent.Left + ColOffset
                    .Top = .Parent.Top + RowOffset
                    .Width = pColWidths(x)
                    Debug.Print x; y; pRstText.Fields(x)
                    If y = 0 Then
                        .Text = pHeadingText(x)
                    Else
                        If Not IsNull(pRstText.Fields(x)) Then .Text = pRstText.Fields(x)
                        .OnAction = pOnAction(x, y - 1)
                    End If
                        
                    If pStyles(x, y - 1) = "" Then
                        If pStyles(1, y - 1) = "" Then
                            .Style = pStylesColl.Find(pStyles(1, 2))
                        Else
                            .Style = pStylesColl.Find(pStyles(1, y))
                        End If
                    Else
                        .Style = pStylesColl.Find(pStyles(x, y - 1))
                    End If
                    
                    .Height = pRowHeight
                End With
                ColOffset = ColOffset + Cell.Width + HPad
        Next
            pRstText.MoveNext
            ColOffset = 0
            RowOffset = RowOffset + Cell.Height + VPad
                    
    Next
    End If
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pCells = New ClsUICells
    Set pStylesColl = New ClsUIStyles
    pCells.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing
    pCells.Terminate
    
    Set pCells = Nothing
    Set pStylesColl = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pCells = Nothing
    Set pStylesColl = Nothing
End Sub
' ===============================================================


