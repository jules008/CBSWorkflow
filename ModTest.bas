Attribute VB_Name = "ModTest"
Option Explicit


Public Sub TestTable()
    Dim MainScreen As ClsUIScreen
    Dim Frame As ClsUIFrame
    Dim SubTable As ClsUITable
    Dim AryColWidths() As Integer
    Dim AryRowHeights() As Integer
    Dim x As Integer
    Dim y As Integer
    Dim AryText() As String
    Dim AryOnAction() As String
    Dim AryStyles() As String
    
    Set Frame = New ClsUIFrame
    Set MainScreen = New ClsUIScreen
    Set SubTable = New ClsUITable
    Set CTimer = New ClsCodeTimer
    
    ReDim AryColWidths(1 To 8)
    ReDim AryRowHeights(1 To 20)
    ReDim AryText(1 To 8, 1 To 20)
    ReDim AryOnAction(1 To 8, 1 To 20)
    ReDim AryStyles(1 To 8, 1 To 20)
    
    CTimer.StartTimer
    
    BuildScreenStyles
    
    With MainScreen
        .Name = "Main Screen"
        .Top = 0
        .Left = 0
        .Width = 1500
        .Height = 1200
        .Style = MENUBAR_STYLE
        .Frames.AddItem Frame, "Frame 1"
        End With
        
    CTimer.MarkTime "Main Screen Built"
    
    AryColWidths(1) = 100
    AryColWidths(2) = 50
    AryColWidths(3) = 50
    AryColWidths(4) = 30
    AryColWidths(5) = 50
    AryColWidths(6) = 50
    AryColWidths(7) = 50
    AryColWidths(8) = 130
    
    AryRowHeights(1) = 30
    AryStyles(1, 1) = "AMBER_CELL"
    AryStyles(1, 2) = "GREEN_CELL"
    
    With SubTable
        .NoCols = 8
        .NoRows = 3
        .ColWidths = AryColWidths
        .StylesColl.Add GREEN_CELL
        .StylesColl.Add AMBER_CELL
        .RowHeights = AryRowHeights
        .Styles = AryStyles
    End With
    
    With Frame
        .Top = 100
        .Left = 100
        .Width = 1200
        .Height = 1000
        .Style = GENERIC_BUTTON
        
        .ReOrder
    End With
    
    CTimer.MarkTime "Frame Built"
    

    AryColWidths(1) = 100
    AryColWidths(2) = 50
    AryColWidths(3) = 50
    AryColWidths(4) = 30
    AryColWidths(5) = 50
    AryColWidths(6) = 50
    AryColWidths(7) = 50
    AryColWidths(8) = 130

    AryRowHeights(1) = 50
'    AryRowHeights(2) = 90

    For x = 1 To 8
        For y = 1 To 20
            AryText(x, y) = x & ", " & y
            AryOnAction(x, y) = "'ModTest.Test'"
            
            If y = 1 Then
                AryStyles(x, y) = "AMBER_CELL"
            Else
                AryStyles(x, y) = "GREEN_CELL"
            End If
        Next
    Next
    

    With Frame.Table
        .Left = 100
        .Top = 100
        .NoCols = 8
        .NoRows = 20
        .HPad = 5
        .VPad = 0
        .SubTableVOff = 50
        .SubTableHOff = 20
        .ColWidths = AryColWidths
        .RowHeights = AryRowHeights
        .OnAction = AryOnAction
        .Text = AryText
        .Styles = AryStyles
        .StylesColl.Add GREEN_CELL
        .StylesColl.Add AMBER_CELL
        .BuildCells 5, 100, SubTable
    End With
    CTimer.MarkTime "Cells Built"
    
    Stop
    Set CTimer = Nothing
    SubTable.Terminate
    Set SubTable = Nothing

    DestroyScreenStyles
    MainScreen.Terminate
End Sub
    
Public Sub DeleteAllShapes()
    Dim Shp As Shape
    
    For Each Shp In ShtMain.Shapes
        Shp.Delete
    Next
End Sub
    
Public Sub Test()
    Dim Frame As ClsUIFrame
    Dim Table As ClsUITable
    
    Set Frame = New ClsUIFrame
    Set Table = New ClsUITable

    Frame.Table = Table
    
    
    Stop


    Set Frame = Nothing
    Set Table = Nothing
End Sub
