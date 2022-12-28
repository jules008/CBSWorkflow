Attribute VB_Name = "ModReport"
'===============================================================
' Module ModReport
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Aug 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModReport"

' ===============================================================
' IntExtCommsReport
' Report of communications due
' ---------------------------------------------------------------
Public Function IntExtCommsReport(RstCommsList As Recordset) As Boolean
    Const NO_COLS As Integer = 5
    Dim Headings(0 To NO_COLS - 1) As String
    Dim Cols(0 To NO_COLS - 1) As Variant
    Dim Align(0 To NO_COLS - 1) As XlHAlign
    Dim ColFormat(0 To NO_COLS - 1) As String
    Dim ReportTitle As String
    Dim AryReport() As Variant
    Dim RstQualDates As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "IntExtCommsReport()"

    On Error GoTo ErrorHandler
    
    With RstCommsList
        If .RecordCount = 0 Then
            MsgBox "There were no results for the report", vbInformation + vbOKOnly
        Else
            .MoveLast
            .MoveFirst
            AryReport = .GetRows(.RecordCount)
        
            ReportTitle = "Internal / External Communication Export"
            Headings(0) = "Contact No"
            Headings(1) = "Name"
            Headings(2) = "Email Address"
            Headings(3) = "Type"
            Headings(4) = "Organisation"
    
            Align(0) = xlHAlignCenter
            Align(1) = xlHAlignLeft
            Align(2) = xlHAlignLeft
            Align(3) = xlHAlignCenter
            Align(4) = xlHAlignLeft
    
            Cols(0) = 15
            Cols(1) = 20
            Cols(2) = 25
            Cols(3) = 15
            Cols(4) = 25
    
            ColFormat(0) = "General"
            ColFormat(1) = "General"
            ColFormat(2) = "General"
            ColFormat(3) = "General"
            ColFormat(4) = "General"
            
            ShtReport.PrintReport AryReport, ReportTitle, Headings, Cols, 4, Align, ColFormat
        End If
    End With
        
GracefulExit:

    IntExtCommsReport = True
    Set RstQualDates = Nothing

Exit Function

ErrorExit:

    Set RstQualDates = Nothing
    IntExtCommsReport = False

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
' Report1
' Report of communications due
' ---------------------------------------------------------------
Public Function Report1(RstReport As Recordset) As Boolean
    Const NO_COLS As Integer = 8
    Dim Headings(0 To NO_COLS - 1) As String
    Dim Cols(0 To NO_COLS - 1) As Variant
    Dim Align(0 To NO_COLS - 1) As XlHAlign
    Dim ColFormat(0 To NO_COLS - 1) As String
    Dim ReportTitle As String
    Dim AryReport() As Variant
    Dim RstQualDates As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Report1()"

    On Error GoTo ErrorHandler
    
    With RstReport
        If .RecordCount = 0 Then
            MsgBox "There were no results for the report", vbInformation + vbOKOnly
        Else
            .MoveLast
            .MoveFirst
            AryReport = .GetRows(.RecordCount)
        
            ReportTitle = "Total Revenue"
            Headings(0) = .Fields(0).Name
            Headings(1) = .Fields(1).Name
            Headings(2) = .Fields(2).Name
            Headings(3) = .Fields(3).Name
            Headings(4) = .Fields(4).Name
            Headings(5) = .Fields(5).Name
            Headings(6) = .Fields(6).Name
            Headings(7) = .Fields(7).Name
    
            Align(0) = xlHAlignLeft
            Align(1) = xlHAlignLeft
            Align(2) = xlHAlignLeft
            Align(3) = xlHAlignCenter
            Align(4) = xlHAlignRight
            Align(5) = xlHAlignRight
            Align(6) = xlHAlignRight
            Align(7) = xlHAlignRight
    
            Cols(0) = 10
            Cols(1) = 15
            Cols(2) = 15
            Cols(3) = 15
            Cols(4) = 15
            Cols(5) = 15
            Cols(6) = 15
            Cols(7) = 15
    
            ColFormat(0) = "General"
            ColFormat(1) = "General"
            ColFormat(2) = "General"
            ColFormat(3) = "General"
            ColFormat(4) = "0.0%"
            ColFormat(5) = "0.0%"
            ColFormat(6) = "£#,##0.00"
            ColFormat(7) = "£#,##0.00"
            
            ShtReport.PrintReport AryReport, ReportTitle, Headings, Cols, 4, Align, ColFormat
        End If
    End With
        
GracefulExit:

    Report1 = True
    Set RstQualDates = Nothing

Exit Function

ErrorExit:

    Set RstQualDates = Nothing
    Report1 = False

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
' Report2
' Report of communications due
' ---------------------------------------------------------------
Public Function Report2(ByVal RstReport As Recordset) As Boolean
    Const NO_COLS As Integer = 3
    Dim Headings(0 To NO_COLS - 1) As String
    Dim Cols(0 To NO_COLS - 1) As Variant
    Dim Align(0 To NO_COLS - 1) As XlHAlign
    Dim ColFormat(0 To NO_COLS - 1) As String
    Dim ReportTitle As String
    Dim AryReport() As Variant
    Dim RstReportData As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Report2()"

    On Error GoTo ErrorHandler
    
    Set RstReportData = RstReport
    
    With RstReportData
        If .RecordCount = 0 Then
            MsgBox "There were no results for the report", vbInformation + vbOKOnly
        Else
            .MoveLast
            .MoveFirst
            
            AryReport = .GetRows(.RecordCount)
            
            For i = LBound(AryReport, 2) To UBound(AryReport, 2)
                AryReport(1, i) = MonthName(AryReport(1, i))
            Next
            
            ReportTitle = "Average Commission Over Time"
            Headings(0) = .Fields(0).Name
            Headings(1) = .Fields(1).Name
            Headings(2) = .Fields(2).Name
    
            Align(0) = xlHAlignLeft
            Align(1) = xlHAlignLeft
            Align(2) = xlHAlignRight
    
            Cols(0) = 15
            Cols(1) = 15
            Cols(2) = 15
    
            ColFormat(0) = "General"
            ColFormat(1) = "General"
            ColFormat(2) = "£#,##0.00"
            
            ShtReport.PrintReport AryReport, ReportTitle, Headings, Cols, 0, Align, ColFormat
        End If
    End With
        
GracefulExit:

    Report2 = True
    Set RstReport = Nothing
    Set RstReportData = Nothing

Exit Function

ErrorExit:

    Set RstReport = Nothing
    Set RstReportData = Nothing
    Report2 = False

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
' Report3
' Report of communications due
' ---------------------------------------------------------------
Public Function Report3(RstReport As Recordset) As Boolean
    Const NO_COLS As Integer = 2
    Dim Headings(0 To NO_COLS - 1) As String
    Dim Cols(0 To NO_COLS - 1) As Variant
    Dim Align(0 To NO_COLS - 1) As XlHAlign
    Dim ColFormat(0 To NO_COLS - 1) As String
    Dim ReportTitle As String
    Dim AryReport() As Variant
    Dim RstQualDates As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Report3()"

    On Error GoTo ErrorHandler
    
    With RstReport
        If .RecordCount = 0 Then
            MsgBox "There were no results for the report", vbInformation + vbOKOnly
        Else
            .MoveLast
            .MoveFirst
            AryReport = .GetRows(.RecordCount)
        
            ReportTitle = "Case Duration (Days)"
            Headings(0) = .Fields(0).Name
            Headings(1) = .Fields(1).Name
    
            Align(0) = xlHAlignLeft
            Align(1) = xlHAlignRight
    
            Cols(0) = 20
            Cols(1) = 15
    
            ColFormat(0) = "General"
            ColFormat(1) = "0.0"
            
            ShtReport.PrintReport AryReport, ReportTitle, Headings, Cols, 2, Align, ColFormat
        End If
    End With
        
GracefulExit:

    Report3 = True
    Set RstQualDates = Nothing

Exit Function

ErrorExit:

    Set RstQualDates = Nothing
    Report3 = False

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
' Report4
' Report of communications due
' ---------------------------------------------------------------
Public Function Report4(RstReport As Recordset) As Boolean
    Const NO_COLS As Integer = 2
    Dim Headings(0 To NO_COLS - 1) As String
    Dim Cols(0 To NO_COLS - 1) As Variant
    Dim Align(0 To NO_COLS - 1) As XlHAlign
    Dim ColFormat(0 To NO_COLS - 1) As String
    Dim ReportTitle As String
    Dim AryReport() As Variant
    Dim RstQualDates As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "Report4()"

    On Error GoTo ErrorHandler
    
    With RstReport
        If .RecordCount = 0 Then
            MsgBox "There were no results for the report", vbInformation + vbOKOnly
        Else
            .MoveLast
            .MoveFirst
            AryReport = .GetRows(.RecordCount)
        
            ReportTitle = "Total Revenue"
            Headings(0) = .Fields(0).Name
            Headings(1) = .Fields(1).Name
    
            Align(0) = xlHAlignLeft
            Align(1) = xlHAlignRight
    
            Cols(0) = 25
            Cols(1) = 15
    
            ColFormat(0) = "General"
            ColFormat(1) = "£#,##0.00"
            
            ShtReport.PrintReport AryReport, ReportTitle, Headings, Cols, 2, Align, ColFormat
        End If
    End With
        
GracefulExit:

    Report4 = True
    Set RstQualDates = Nothing

Exit Function

ErrorExit:

    Set RstQualDates = Nothing
    Report4 = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function




