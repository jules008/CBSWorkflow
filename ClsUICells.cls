VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUICells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUICells
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
Private pCells As Collection
Private pParent As ClsUITable

'---------------------------------------------------------------
Public Property Get Parent() As ClsUITable
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUITable)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCells.[_NewEnum]
End Function

' ===============================================================
' Method FindCell
' Finds Cell from index
'---------------------------------------------------------------
Public Function FindCell(Name As Variant) As ClsUICell
Attribute FindCell.VB_UserMemId = 0
    On Error Resume Next
    Set FindCell = pCells.Item(Name)
End Function

' ===============================================================
' Method AddCell
' Add Cell to Collection
'---------------------------------------------------------------
Public Sub AddCell(UICell As ClsUICell, ByVal Col As Integer, ByVal Row As Integer)
    UICell.Initialize Me.Parent
    UICell.Col = Col
    UICell.Row = Row
    UICell.Name = Col & "-" & Row
    pCells.Add UICell, Key:=UICell.Name
End Sub

' ===============================================================
' Method RemoveCell
' Deletes Cell from collection
'---------------------------------------------------------------
Public Sub RemoveCell(Name As Variant)
    pCells.Remove Name
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all Cells from collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Cell As ClsUICell
    
    Sin 0
    For Each Cell In pCells
        pCells.Remove Cell.Name
        Cell.Terminate
        Set Cell = Nothing
    Next
End Sub

' ===============================================================
' Method CellCount
' Returns number of Cells in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pCells.Count
End Function

' ===============================================================
' ===============================================================
Public Sub Initialize(obj As ClsUITable)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pCells = New Collection
End Sub

' ===============================================================
Public Sub Terminate()
    Dim Cell As ClsUICell
    
    For Each Cell In pCells
        Cell.Terminate
    Next
    
    Set pParent = Nothing
    Set pCells = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pCells = Nothing
End Sub

'---------------------------------------------------------------
