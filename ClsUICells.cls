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
Private pUICells As Collection
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
    Set NewEnum = pUICells.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------
Public Function FindItem(Name As Variant) As ClsUICell
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUICells.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(UICell As ClsUICell)
    UICell.Initialize Me.Parent
    pUICells.Add UICell, Key:=UICell.Row & "-" & UICell.Col
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(Name As Variant)
    pUICells.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pUICells.Count
End Function

' ===============================================================
' Method Text
' Adds text to the line item
'---------------------------------------------------------------
Public Sub Text(Row As Integer, Col As Integer, ItemText As String, LineStyle As TypeStyle, Optional OnAction As String, Optional Badge As Shape)
    Dim UICell As ClsUICell
    Dim LocCustomStyle As TypeStyle
    Dim LeftOffset As Integer
    Dim i As Integer
    
    Do While i <> Col
        LeftOffset = LeftOffset + pColumns(i)
        i = i + 1
    Loop
    
    Set UICell = New ClsUICell
    Me.AddItem UICell
    
    With UICell
        .Style = LineStyle
        .Left = Left + LeftOffset
        .Width = pColumns(Col)
        .Height = pHeight
        .Top = Top + (Row * pRowOffset)
        If Not Badge Is Nothing Then .Badge = Badge
        .Text = ItemText
        .OnAction = OnAction
    End With
    
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the Cells sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Cell As ClsUICell
    For Each Cell In pUICells
        Cell.ShpCell.ZOrder msoBringToFront
        Cell.ReOrder
    Next
    
End Sub


' ===============================================================
Public Sub Initialize(obj As ClsUITable)
    Set Me.Parent = obj
End Sub

Private Sub Class_Initialize()
    Set pUICells = New Collection
End Sub

' ===============================================================
Public Sub Terminate()
    Dim UICell As ClsUICell
    
    For Each UICell In pUICells
        UICell.Terminate
    Next
    
    Set pParent = Nothing
    Set pUICells = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUICells = Nothing
End Sub

'---------------------------------------------------------------
