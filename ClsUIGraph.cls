VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIGraph
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Dec 20
'===============================================================
' Methods
'---------------------------------------------------------------
' GenGraph - Generates doughnut Grpah
'===============================================================

Option Explicit
Private pName As String
Private pChartType As EnChartType
Private pDataLabels As Boolean
Private pHeight As Integer
Private pWidth As Integer
Private pLeft As Integer
Private pTop As Integer
Private pTxtMainLbl As String
Private pTxtSmallLbl As String
Private pTitle As String
Private pSer1Colour As Long
Private pSer2Colour As Long
Private pSer3Colour As Long
Private pSer4Colour As Long
Private pSer5Colour As Long
Private pSer6Colour As Long
Private pSer1Name As String
Private pSer2Name As String
Private pSer3Name As String
Private pBackColour As Long
Private pSourceData() As Variant
Private pCht As Chart
Private pTxtTitle As Shape
Private pTxtPC As Shape
Private pTxtLabel As Shape
Private pVisible As Boolean
Private pParent As ClsUIFrame

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ChartType() As EnChartType
    ChartType = pChartType
End Property

Public Property Let ChartType(ByVal vNewValue As EnChartType)
    pChartType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DataLabels() As Boolean
    DataLabels = pDataLabels
End Property

Public Property Let DataLabels(ByVal vNewValue As Boolean)
    pDataLabels = vNewValue
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
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TxtMainLbl() As String
    TxtMainLbl = pTxtMainLbl
End Property

Public Property Let TxtMainLbl(ByVal vNewValue As String)
    pTxtMainLbl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TxtSmallLbl() As String
    TxtSmallLbl = pTxtSmallLbl
End Property

Public Property Let TxtSmallLbl(ByVal vNewValue As String)
    pTxtSmallLbl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Title() As String
    Title = pTitle
End Property

Public Property Let Title(ByVal vNewValue As String)
    pTitle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser1Name(ByVal vNewValue As String)
    pSer1Name = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser2Name(ByVal vNewValue As String)
    pSer2Name = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser3Name(ByVal vNewValue As String)
    pSer3Name = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser1Colour(ByVal vNewValue As Long)
    pSer1Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser2Colour(ByVal vNewValue As Long)
    pSer2Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser3Colour(ByVal vNewValue As Long)
    pSer3Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser4Colour(ByVal vNewValue As Long)
    pSer4Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser5Colour(ByVal vNewValue As Long)
    pSer5Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Let Ser6Colour(ByVal vNewValue As Long)
    pSer6Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get BackColour() As Long
    BackColour = pBackColour
End Property

Public Property Let BackColour(ByVal vNewValue As Long)
    pBackColour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SourceData() As Variant()
    SourceData() = pSourceData()
End Property

Public Property Let SourceData(ByRef vNewValue() As Variant)
    pSourceData() = vNewValue()
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
        
    If pVisible Then
        Select Case pChartType
            Case enDoNut
                pTxtTitle.Visible = msoTrue
                pTxtPC.Visible = msoCTrue
                pTxtLabel.Visible = msoCTrue
            Case Else
                pTxtTitle.Visible = msoTrue
                pTxtPC.Visible = msoFalse
                pTxtLabel.Visible = msoFalse
        End Select
    Else
        pTxtTitle.Visible = msoFalse
        pTxtPC.Visible = msoFalse
        pTxtLabel.Visible = msoFalse
    End If
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsUIFrame)
    Set pParent = obj
End Property

' ===============================================================
' Method GenGraph
' Generates doughnut Grpah
'---------------------------------------------------------------
Public Sub GenGraph()
    Select Case pChartType
        Case enDoNut
            GenDoNut
        Case enBarStacked
            GenBarStacked
        Case enBarHoriz
            GenBarHoriz
        Case enline
            GenLine
        Case enPie
            GenPie
    End Select
End Sub

' ===============================================================
' Method GenDoNut
' Generates doughnut Grpah
'---------------------------------------------------------------
Private Sub GenDoNut()
    
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 1.75
            .Format.Fill.ForeColor.RGB = pBackColour
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlDoughnut
        .ChartTitle.Text = pTitle
        .ChartTitle.IncludeInLayout = True
        .HasLegend = True
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = "strName"
            .Values = GetArrayRow(pSourceData, 1)
            .Points(1).Format.Fill.ForeColor.RGB = pSer1Colour
            .Points(2).Format.Fill.ForeColor.RGB = pSer2Colour
            .Points(3).Format.Fill.ForeColor.RGB = pSer3Colour
             .XValues = GetArrayRow(pSourceData, 0)
        End With
        
        With .ChartGroups(1)
            .DoughnutHoleSize = 70
        End With
        
    End With

End Sub

' Method GenPie
' Generates doughnut Grpah
'---------------------------------------------------------------
Private Sub GenPie()
    Dim Subtotal As Single
    Dim Total As Single
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 1.75
            .Format.Fill.ForeColor.RGB = pBackColour
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlPie
        .ChartTitle.Text = pTitle
        .ChartTitle.IncludeInLayout = True
        .HasLegend = False
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64"), xlColumns
        
        With .SeriesCollection(1)
            .Name = "strName"
            .Values = GetArrayRow(pSourceData, 1)
            .Points(1).Format.Fill.ForeColor.RGB = pSer1Colour
            .Points(2).Format.Fill.ForeColor.RGB = pSer2Colour
            .Points(3).Format.Fill.ForeColor.RGB = pSer3Colour
            .Points(4).Format.Fill.ForeColor.RGB = pSer4Colour
            .Points(5).Format.Fill.ForeColor.RGB = pSer5Colour
            .Points(6).Format.Fill.ForeColor.RGB = pSer6Colour
            .XValues = GetArrayRow(pSourceData, 0)
        End With
        
        .ApplyDataLabels
        
        With .SeriesCollection(1).DataLabels
            .ShowCategoryName = True
            .ShowValue = False
            .Position = xlLabelPositionBestFit
            .Position = xlLabelPositionOutsideEnd
            '.Position = xlLabelPositionInsideEnd
            End With
        End With
End Sub

' ===============================================================
' Method GenBarStacked
' Generates Stacked bar Grpah
'---------------------------------------------------------------
Private Sub GenBarStacked()
    Dim XValues() As Variant
    
    XValues = GetArrayRow(pSourceData, 0)
    
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 1.7
            .Format.Fill.ForeColor.RGB = COL_WHITE
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlColumnStacked
        .ChartTitle.Text = pTitle
        .HasLegend = True
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = pSer1Name
            .Values = GetArrayRow(pSourceData, 1)
            .Format.Fill.ForeColor.RGB = pSer1Colour
            
            If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            .XValues = XValues
        End With
                
        If UBound(pSourceData, 1) > 1 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                .Name = pSer2Name
                .Values = GetArrayRow(pSourceData, 2)
                .Format.Fill.ForeColor.RGB = pSer2Colour
                If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            End With
        End If
        
        If UBound(pSourceData, 1) > 2 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(3)
                .Name = pSer2Name
                .Values = GetArrayRow(pSourceData, 3)
                .Format.Fill.ForeColor.RGB = pSer3Colour
                If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            End With
        End If
        
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm yy"
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory).ReversePlotOrder = False
        
        With .ChartGroups(1)
            .GapWidth = 80
        End With
        
        With .PlotArea
            .Format.Fill.Transparency = 1
            .Height = pCht.ChartArea.Height / 100 * 85
            .Width = pCht.ChartArea.Width / 100 * 80
            .Top = 30
            .Left = 5
        End With
    
    End With

End Sub

' ===============================================================
' Method GenBarHoriz
' Generates horizontal bar Grpah
'---------------------------------------------------------------
Private Sub GenBarHoriz()
    Dim XValues() As Variant
    Dim DataLabSub() As Variant
    Dim DataLabTot() As Variant
    Dim i As Integer
    
    XValues = GetArrayCol(pSourceData, 1)
    DataLabSub = GetArrayCol(pSourceData, 2)
    DataLabTot = GetArrayCol(pSourceData, 3)
    
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 1.7
            .Format.Fill.ForeColor.RGB = COL_BLACK
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlBarClustered
        .ChartTitle.Text = pTitle
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = pSourceData(1, 2)
            .Values = GetArrayCol(pSourceData, 4)
            .Format.Fill.ForeColor.RGB = pSer1Colour
            .XValues = XValues
            
            If pDataLabels Then
                .ApplyDataLabels
                For i = LBound(DataLabTot) To UBound(DataLabTot)
                    .Points(i).DataLabel.Text = DataLabSub(i) & " of " & DataLabTot(i)
                    .Points(i).DataLabel.Position = xlLabelPositionInsideEnd
                Next
            End If
        End With
        
        With .Axes(xlValue)
            .TickLabels.NumberFormat = "0%"
            .MaximumScale = 1
        End With
        
        With .ChartGroups(1)
            .GapWidth = 80
        End With
        
        With .PlotArea
            .Format.Fill.Transparency = 1
            .Height = pCht.ChartArea.Height / 100 * 85
            .Width = pCht.ChartArea.Width / 100 * 100
            .Top = 30
            .Left = 5
        End With
    
    End With

End Sub


' ===============================================================
' Method GenLine
' Generates Line Grpah
'---------------------------------------------------------------
Private Sub GenLine()
    Dim XValues() As Variant
    Dim i As Integer
        
    pCht.Parent.Name = pName
    XValues = GetArrayRow(pSourceData, 0)
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 2
            .Format.Fill.ForeColor.RGB = COL_WHITE
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlLine
        .ChartTitle.Text = pTitle
        .HasLegend = True
        .ChartTitle.IncludeInLayout = False
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = pCht.ChartArea.Width / 100 * 4
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = pSer1Name
            .Values = GetArrayRow(pSourceData, 1)
            .Format.Fill.ForeColor.RGB = pSer1Colour
            
            If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            .XValues = XValues
        End With
        
        If UBound(pSourceData, 1) > 1 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                .Name = pSer2Name
                .Values = GetArrayRow(pSourceData, 2)
                .Format.Line.ForeColor.RGB = pSer2Colour
                If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
        End With
        End If
        
        If UBound(pSourceData, 1) > 2 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(3)
                .Name = pSer3Name
                .Values = GetArrayRow(pSourceData, 3)
                .Format.Line.ForeColor.RGB = pSer3Colour
                If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            End With
        End If
        
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm yy"
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory).ReversePlotOrder = False
        
        With .PlotArea
            .Format.Line.Transparency = 1
            .Height = pCht.ChartArea.Height / 100 * 85
            .Width = pCht.ChartArea.Width / 100 * 80
            .Top = 30
            .Left = 5
        End With
    
    End With

End Sub

' ===============================================================
' Method DestroyGraph
' destroys all graph cells
'---------------------------------------------------------------
Public Sub DestroyGraph()
    pCht.Parent.Delete
    pTxtTitle.Delete
    pTxtPC.Delete
    pTxtLabel.Delete
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    ShtMain.Shapes(pName).ZOrder msoBringToFront
    pTxtTitle.ZOrder msoBringToFront
    pTxtLabel.ZOrder msoBringToFront
    pTxtPC.ZOrder msoBringToFront
End Sub

' ===============================================================
' Method GetArrayRow
' Gets a single row of a two dimensional array
'---------------------------------------------------------------
Private Function GetArrayRow(ArySource() As Variant, RowNo As Integer) As Variant()
    Dim i As Integer
    Dim AryOutput As Variant
    
    ReDim AryOutput(LBound(ArySource, 2) To UBound(ArySource, 2))
    
    For i = LBound(ArySource, 2) To UBound(ArySource, 2)
        If IsNull(ArySource(RowNo, i)) Then
            AryOutput(i) = 0
        Else
        AryOutput(i) = ArySource(RowNo, i)
        End If
    Next
    GetArrayRow = AryOutput
End Function
' ===============================================================
' Method GetArrayCol
' Gets a single Col of a two dimensional array
'---------------------------------------------------------------
Private Function GetArrayCol(ArySource() As Variant, ColNo As Integer) As Variant()
    Dim i As Integer
    Dim AryOutput As Variant
    
    ReDim AryOutput(LBound(ArySource, 1) To UBound(ArySource, 1) - 1)
    
    For i = LBound(ArySource, 1) + 1 To UBound(ArySource, 1)
        AryOutput(i - 1) = ArySource(i, ColNo)
    Next
    GetArrayCol = AryOutput
End Function

' ===============================================================
Public Sub Initialize(obj As ClsUIFrame)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------
Private Sub Class_Initialize()
    Set pCht = ShtMain.Shapes.AddChart2.Chart
    Set pTxtPC = ShtMain.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
    Set pTxtLabel = ShtMain.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
    Set pTxtTitle = ShtMain.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
End Sub
' ===============================================================
Public Sub Terminate()
    Set pParent = Nothing
    
    pCht.Parent.Delete
    pTxtTitle.Delete
    pTxtPC.Delete
    pTxtLabel.Delete
    
    Set pCht = Nothing
    Set pTxtPC = Nothing
    Set pTxtLabel = Nothing
    Set pTxtTitle = Nothing
End Sub
'---------------------------------------------------------------

Private Sub Class_Terminate()
    Set pCht = Nothing
    Set pTxtPC = Nothing
    Set pTxtLabel = Nothing
    Set pTxtTitle = Nothing
End Sub
' ===============================================================

