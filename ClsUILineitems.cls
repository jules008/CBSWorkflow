VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUILineitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class Lineitems
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
Private pNoColumns As Integer
Private pColumns() As Integer
Private pUILineitems As Collection
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pStyle As TypeStyle
Private pRowOffset As Integer
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get NoColumns() As Integer
    NoColumns = pNoColumns
End Property

Public Property Let NoColumns(ByVal vNewValue As Integer)
    pNoColumns = vNewValue
    ReDim pColumns(0 To NoColumns - 1)
End Property

'---------------------------------------------------------------
Public Property Get ColumnWidth(i As Integer) As Integer
    ColumnWidth = pColumns(i)
End Property

Public Property Let Columns(vNewValue As String)
    Dim i As Integer
    Dim ColArray() As String
    
    ColArray() = Split(vNewValue, ":")
    
    For i = 0 To pNoColumns - 1
        pColumns(i) = CInt(ColArray(i))
    Next
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue + Parent.Top
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue + Parent.Left
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pHeight
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pWidth
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RowOffset() As Integer
    RowOffset = pRowOffset
End Property

Public Property Let RowOffset(ByVal vNewValue As Integer)
    pRowOffset = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUILineitems.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUILineitem
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUILineitems.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UILineitem As ClsUILineitem)
    If UILineitem Is Nothing Then
        Set UILineitem = New ClsUILineitem
    End If
    UILineitem.Initialize Me.Parent
    pUILineitems.Add UILineitem, Key:=UILineitem.Name
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUILineitems.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUILineitems.Count
End Function

' ===============================================================
' Method Text
' Adds text to the line item
'---------------------------------------------------------------
Public Sub Text(Row As Integer, Col As Integer, ItemText As String, LineStyle As TypeStyle, Optional OnAction As String, Optional Badge As Shape)
    Dim UILineitem As ClsUILineitem
    Dim LocCustomStyle As TypeStyle
    Dim LeftOffset As Integer
    Dim i As Integer
    
    Do While i <> Col
        LeftOffset = LeftOffset + pColumns(i)
        i = i + 1
    Loop
    
    Set UILineitem = New ClsUILineitem
    Me.AddItem UILineitem
    
    With UILineitem
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
' re-orders the shapes so that the Lineitems sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Lineitem As ClsUILineitem
    For Each Lineitem In pUILineitems
        Lineitem.ShpLineitem.ZOrder msoBringToFront
        Lineitem.ReOrder
    Next
    
End Sub


' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
    Set pUILineitems = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    
End Sub

'---------------------------------------------------------------
