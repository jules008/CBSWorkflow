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

