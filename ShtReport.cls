VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtReport
'===============================================================
' v1.0.0 - Initial Version
' v1.0.1 - Added sort to report
'---------------------------------------------------------------
' Date - 19 Mar 20
'===============================================================
Option Explicit

Private Const StrMODULE As String = "ShtReport"
Private Const REPORT_RNG As String = "A:AE"
Private Const RNG_IND As String = "A1"
Private Const HEADING As String = "A3"
Private Const REP_IND As String = "A4"

' ===============================================================
' PrintReport
' sets report data for printing
' ---------------------------------------------------------------
Public Function PrintReport(AryReport() As Variant, TxtTitle As String, Headings() As String, Cols() As Variant, SortByCol As Integer, Alignment() As XlHAlign, ColFormat() As String) As Boolean
    Dim i As Integer
    Dim AryTransposed() As Variant
    Dim RngReport As Range
    Dim FileName As String
    
    Const StrPROCEDURE As String = "PrintReport()"

    On Error GoTo ErrorHandler
    
    ClearReport
    
    AryTransposed = ModLibrary.TransposeArray(AryReport)
    
    For i = LBound(Headings) To UBound(Headings)
        Range(HEADING).Offset(0, i) = Headings(i)
        Range(HEADING).Offset(0, i).ColumnWidth = Cols(i)
        Range(HEADING).Offset(0, i).EntireColumn.HorizontalAlignment = Alignment(i)
        Range(HEADING).Offset(0, i).EntireColumn.NumberFormat = ColFormat(i)
    Next
    
    Range(RNG_IND).HorizontalAlignment = xlHAlignLeft
    Range(RNG_IND) = TxtTitle
    
    PerfSettingsOn
    Me.Visible = xlSheetVisible
    
    Set RngReport = Range(REP_IND).Resize(UBound(AryTransposed, 1) + 1, UBound(AryTransposed, 2) + 1)
    
    RngReport = AryTransposed
    
    Me.PageSetup.LeftMargin = 15
    Me.PageSetup.RightMargin = 15
        
    If SortByCol <> 0 Then RngReport.Sort key1:=Range(REP_IND).Offset(0, SortByCol - 1)
    
    FileName = Int(2 + Rnd * (30 - 2 + 1))
    
    Me.Copy
    
    PerfSettingsOff
    
    Me.Visible = xlSheetVeryHidden
    
    PrintReport = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    PrintReport = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ClearReport
' Clears data from report sheet
' ---------------------------------------------------------------
Public Sub ClearReport()
    Range(REPORT_RNG).ClearContents
End Sub

' ===============================================================
' DataExport
' Creates Report from Recordset
' ---------------------------------------------------------------
Public Function DataExport(RstData As Recordset, ColWidths() As Integer, Headings() As String, ColFormats() As String) As Boolean
    Dim ReportBook As Workbook
    Dim RngQry As Range
    Dim RngHeader As Range
    Dim ShtReport As Worksheet
    Dim i As Integer
    
    Const StrPROCEDURE As String = "DataExport()"

    On Error GoTo ErrorHandler

    Set ReportBook = Workbooks.Add
    Set ShtReport = ReportBook.Worksheets(1)
    
    With ShtReport
        Set RngQry = .Range(RNG_IND)
        
        'headings and col widths
        For i = 0 To UBound(Headings)
            RngQry.Offset(0, i) = Headings(i)
            RngQry.Offset(0, i).ColumnWidth = ColWidths(i)
        Next
        
        'formats
        For i = 0 To UBound(ColFormats)
            .Columns(i + 1).NumberFormat = ColFormats(i)
        Next
        
        'format heading
        Set RngHeader = .Range(.Cells(1, 1), .Cells(1, UBound(Headings) + 1))
    
        With RngHeader
            .Interior.Color = COL_BLUE
            .Borders.Color = COL_GREEN
            .Font.Bold = True
        
            'set filter
            .AutoFilter
        End With
    
        RngQry.Offset(1, 0).CopyFromRecordset RstData
    End With
    
    Set RngQry = Nothing
    Set RngHeader = Nothing
    Set ReportBook = Nothing
    Set ShtReport = Nothing
    DataExport = True

Exit Function

ErrorExit:

    Set RngQry = Nothing
    Set RngHeader = Nothing
    Set ReportBook = Nothing
    Set ShtReport = Nothing
    
'    ***CleanUpCode***
    DataExport = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


