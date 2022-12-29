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
Public Property Get Ser1Colour() As Long
    Ser1Colour = pSer1Colour
End Property

Public Property Let Ser1Colour(ByVal vNewValue As Long)
    pSer1Colour = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Ser2Colour() As Long
    Ser2Colour = pSer2Colour
End Property

Public Property Let Ser2Colour(ByVal vNewValue As Long)
    pSer2Colour = vNewValue
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
    Dim Subtotal As Single
    Dim Total As Single
    
    Subtotal = pSourceData(1)
    Total = pSourceData(2)
    
    pSourceData(2) = Total - Subtotal
    
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height
            .Format.Fill.ForeColor.RGB = pBackColour
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlDoughnut
        .ChartTitle.Text = ""
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = "strName"
            .Values = pSourceData
            .Points(1).Format.Fill.ForeColor.RGB = pSer1Colour
            .Points(2).Format.Fill.ForeColor.RGB = pSer2Colour
        End With
        
        With .ChartGroups(1)
            .DoughnutHoleSize = 80
        End With
        
        With .PlotArea
            .Format.Fill.Transparency = 1
            .Top = pCht.ChartArea.Height / 100 * 5
            .Height = pCht.ChartArea.Height / 100 * 70
            .Left = (pCht.ChartArea.Width - .Width) / 2
        End With
    
    End With
    
    With pTxtTitle
        .Name = "pTxtTitle"
        .Height = pCht.ChartArea.Height / 100 * 25
        .Width = pCht.ChartArea.Width / 100 * 120
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + pCht.ChartArea.Height / 100 * 75
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtTitle.Height / 100 * 40
                .Font.Color = RGB(0, 0, 0)
                .Text = pTitle
            End With
        End With
    End With
    
    With pTxtPC
        .Name = "pTxtPC"
        .Height = pCht.ChartArea.Height / 100 * 25
        .Width = pCht.ChartArea.Width / 100 * 50
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + (pCht.ChartArea.Height / 2) - (.Height / 2) - (pCht.ChartArea.Height / 100 * 10)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtPC.Height / 100 * 80
                .Font.Color = RGB(0, 0, 0)
                .Text = Format(Subtotal / Total, "0%")
            End With
        End With
    End With
    
    With pTxtLabel
        .Name = "TxtLabel"
        .Height = pCht.ChartArea.Height / 100 * 5
        .Width = pCht.ChartArea.Width / 100 * 50
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + (pCht.ChartArea.Height / 2) - (.Height / 2) + (pCht.ChartArea.Height / 100 * 7)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtPC.Height / 100 * 25
                .Font.Color = RGB(0, 0, 0)
                .Text = Subtotal & " of " & Total
            End With
        End With
    End With

End Sub

' Method GenPie
' Generates doughnut Grpah
'---------------------------------------------------------------
Private Sub GenPie()
    pCht.Parent.Name = pName
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height
            .Format.Fill.ForeColor.RGB = pBackColour
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlDoughnut
        .ChartTitle.Text = ""
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = "strName"
            .Values = pSourceData
            .Points(1).Format.Fill.ForeColor.RGB = pSer1Colour
            .Points(2).Format.Fill.ForeColor.RGB = pSer2Colour
        End With
        
        With .ChartGroups(1)
            .DoughnutHoleSize = 80
        End With
        
        With .PlotArea
            .Format.Fill.Transparency = 1
            .Top = pCht.ChartArea.Height / 100 * 5
            .Height = pCht.ChartArea.Height / 100 * 70
            .Left = (pCht.ChartArea.Width - .Width) / 2
        End With
    
    End With
    
    With pTxtTitle
        .Name = "pTxtTitle"
        .Height = pCht.ChartArea.Height / 100 * 25
        .Width = pCht.ChartArea.Width / 100 * 120
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + pCht.ChartArea.Height / 100 * 75
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtTitle.Height / 100 * 40
                .Font.Color = RGB(0, 0, 0)
                .Text = pTitle
            End With
        End With
    End With
    
    With pTxtPC
        .Name = "pTxtPC"
        .Height = pCht.ChartArea.Height / 100 * 25
        .Width = pCht.ChartArea.Width / 100 * 50
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + (pCht.ChartArea.Height / 2) - (.Height / 2) - (pCht.ChartArea.Height / 100 * 10)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtPC.Height / 100 * 80
                .Font.Color = RGB(0, 0, 0)
                .Text = Format(Subtotal / Total, "0%")
            End With
        End With
    End With
    
    With pTxtLabel
        .Name = "TxtLabel"
        .Height = pCht.ChartArea.Height / 100 * 5
        .Width = pCht.ChartArea.Width / 100 * 50
        .Left = pCht.ChartArea.Left + (pCht.ChartArea.Width / 2) - (.Width / 2) + 3
        .Top = pCht.ChartArea.Top + (pCht.ChartArea.Height / 2) - (.Height / 2) + (pCht.ChartArea.Height / 100 * 7)
        .Fill.Transparency = 1
        .Line.Visible = msoFalse
        
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            With .Characters
                .Font.FontStyle = "Calibri"
                .Font.Size = pTxtPC.Height / 100 * 25
                .Font.Color = RGB(0, 0, 0)
                .Text = Subtotal & " of " & Total
            End With
        End With
    End With

End Sub

' ===============================================================
' Method GenBarStacked
' Generates Stacked bar Grpah
'---------------------------------------------------------------
Private Sub GenBarStacked()
    Dim XValues() As Variant
    
    XValues = GetArrayRow(pSourceData, 1)
    
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
        
        .ChartType = xlColumnStacked
        .ChartTitle.Text = pTitle
        .HasLegend = True
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = pCht.ChartArea.Width / 100 * 4
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = pSourceData(2, 1)
            .Values = GetArrayRow(pSourceData, 2)
            .Format.Fill.ForeColor.RGB = pSer1Colour
            
            If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            .XValues = XValues
        End With
                
        If UBound(pSourceData, 1) > 2 Then
            .SeriesCollection.NewSeries
            With .SeriesCollection(2)
                .Name = pSourceData(3, 1)
                .Values = GetArrayRow(pSourceData, 3)
                .Format.Fill.ForeColor.RGB = pSer2Colour
                If pDataLabels Then .ApplyDataLabels xlDataLabelsShowValue
            End With
        End If
        
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm yy"
        
        With .ChartGroups(1)
            .GapWidth = 117
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
            .Font.Size = pCht.ChartArea.Width / 100 * 4
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
    XValues = GetArrayRow(pSourceData, 1)
    
    With pCht
        With .ChartArea
            .Left = Parent.Left + pLeft
            .Top = Parent.Top + pTop
            .Height = pHeight
            .Width = .Height * 1.7
            .Format.Fill.ForeColor.RGB = COL_BLACK
            .Border.LineStyle = xlNone
        End With
        
        .ChartType = xlLine
        .ChartTitle.Text = pTitle
        
        With .ChartTitle.Format.TextFrame2.TextRange
            .Font.Name = "Calibri"
            .Font.Size = pCht.ChartArea.Width / 100 * 4
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
        
        .SetSourceData ShtMain.Range("H64:I64")
        
        With .SeriesCollection(1)
            .Name = pSourceData(1, 2)
            .Values = GetArrayRow(pSourceData, 2)
            .Format.Fill.ForeColor.RGB = pSer1Colour
            .XValues = XValues
        End With
        
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "0%"
            .ReversePlotOrder = True
            .Crosses = xlMaximum
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
    
    ReDim AryOutput(LBound(ArySource, 2) To UBound(ArySource, 2) - 1)
    
    For i = LBound(ArySource, 2) + 1 To UBound(ArySource, 2)
        AryOutput(i - 1) = ArySource(RowNo, i)
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

