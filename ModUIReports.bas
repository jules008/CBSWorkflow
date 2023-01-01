Attribute VB_Name = "ModUIReports"
'===============================================================
' Module ModUIReports
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Dec 22
'===============================================================							 				  

Option Explicit

Private Const StrMODULE As String = "ModUIReports"

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ShtMain.Unprotect PROTECT_KEY
    
    Application.ScreenUpdating = False
    
    If Not BuildMainFrame1 Then Err.Raise HANDLED_ERROR
    If Not BuildMainFrame2 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn1 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn2 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn3 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn4 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn5 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn6 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn7 Then Err.Raise HANDLED_ERROR
    If Not BuildScreenBtn8 Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
           
    Application.ScreenUpdating = True
    
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
    
    ModLibrary.PerfSettingsOff
                    
    BuildScreen = True
       
Exit Function

ErrorExit:
    
    Application.ScreenUpdating = True
    
    ModLibrary.PerfSettingsOff

    BuildScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainFrame1
' ---------------------------------------------------------------
Private Function BuildMainFrame1() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame1()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
        .Top = REP_FRAME_TOP
        .Left = REP_FRAME_LEFT
        .Width = REP_FRAME_WIDTH
        .Height = REP_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
		.ZOrder = 1
        
        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
            .Text = "Reports"
            .Style = HEADER_STYLE
            .Visible = True
        End With
        
    End With
    
    BuildMainFrame1 = True

Exit Function

ErrorExit:

    BuildMainFrame1 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainFrame2
' ---------------------------------------------------------------
Private Function BuildMainFrame2() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame2()"

    On Error GoTo ErrorHandler

    Set MainFrame2 = New ClsUIFrame
    MainScreen.Frames.AddItem MainFrame2, "Main Frame 2"
    
    'add main frame
    With MainFrame2
        .Top = REP_FRAME_2_TOP
        .Left = REP_FRAME_2_LEFT
        .Width = REP_FRAME_2_WIDTH
        .Height = REP_FRAME_2_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .ZOrder = 1
        
        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame 2 Header"
            .Text = "Data Exports"
            .Style = HEADER_STYLE
            .Visible = True
         End With
        
    End With
    
    BuildMainFrame2 = True

Exit Function

ErrorExit:

    BuildMainFrame2 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn1
' ---------------------------------------------------------------
Private Function BuildScreenBtn1() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn1()"

    On Error GoTo ErrorHandler

    Set BtnReport1 = New ClsUIButton
    
    With BtnReport1
        .Name = "BtnRep1"
        MainFrame.Buttons.Add BtnReport1
        .Height = BTN_REP_1_HEIGHT
        .Left = BTN_REP_1_LEFT
        .Top = BTN_REP_1_TOP
        .Width = BTN_REP_1_WIDTH
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & 0 & ":" & enBtnReport & ":1" & """)'"
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Total Revenue" & vbCr & vbCr & "Revenue based on the commission and exit fee income"
        .Visible = True
    End With
    
    
    BuildScreenBtn1 = True

Exit Function

ErrorExit:

    BuildScreenBtn1 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn2
' ---------------------------------------------------------------
Private Function BuildScreenBtn2() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn2()"

    On Error GoTo ErrorHandler

    Set BtnReport2 = New ClsUIButton
    
    With BtnReport2
        .Name = "BtnRep2"
        MainFrame.Buttons.Add BtnReport2
        .Height = BTN_REP_2_HEIGHT
        .Left = BTN_REP_2_LEFT
        .Top = BTN_REP_2_TOP
        .Width = BTN_REP_2_WIDTH
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & 0 & ":" & enBtnReport & ":2" & """)'"
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Average Commission Over Time" & vbCr & vbCr & "Monthly trend of average Commision earned per case"
    End With
    
    
    BuildScreenBtn2 = True

Exit Function

ErrorExit:

    BuildScreenBtn2 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn3
' ---------------------------------------------------------------
Private Function BuildScreenBtn3() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn3()"

    On Error GoTo ErrorHandler

    Set BtnReport3 = New ClsUIButton
    
    With BtnReport3
        .Name = "BtnRep3"
        MainFrame.Buttons.Add BtnReport3
        .Height = BTN_REP_3_HEIGHT
        .Left = BTN_REP_3_LEFT
        .Top = BTN_REP_3_TOP
        .Width = BTN_REP_3_WIDTH
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & 0 & ":" & enBtnReport & ":3" & """)'"
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Case Duration" & vbCr & vbCr & "Average case duration for each Case Manager"
    End With
    
    
    BuildScreenBtn3 = True

Exit Function

ErrorExit:

    BuildScreenBtn3 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildScreenBtn4
' ---------------------------------------------------------------
Private Function BuildScreenBtn4() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn4()"

    On Error GoTo ErrorHandler

    Set BtnReport4 = New ClsUIButton
    
    With BtnReport4
        .Name = "BtnRep4"
        MainFrame.Buttons.Add BtnReport4
        .Height = BTN_REP_4_HEIGHT
        .Left = BTN_REP_4_LEFT
        .Top = BTN_REP_4_TOP
        .Width = BTN_REP_4_WIDTH
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & 0 & ":" & enBtnReport & ":4" & """)'"
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Debt Per Client" & vbCr & vbCr & "Total debt currently being acquired for clients"
    End With
    
    
    BuildScreenBtn4 = True

Exit Function

ErrorExit:

    BuildScreenBtn4 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn5
' ---------------------------------------------------------------
Private Function BuildScreenBtn5() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn5()"

    On Error GoTo ErrorHandler

    Set BtnReport5 = New ClsUIButton
    
    With BtnReport5
        .Name = "BtnExpt5"
        MainFrame2.Buttons.Add BtnReport5
        .Height = BTN_EXP_5_HEIGHT
        .Left = BTN_EXP_5_LEFT
        .Top = BTN_EXP_5_TOP
        .Width = BTN_EXP_5_WIDTH
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Export 5" & vbCr & vbCr & "Export 5 description"
        .Visible = True
    End With
    
    
    BuildScreenBtn5 = True

Exit Function

ErrorExit:

    BuildScreenBtn5 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn6
' ---------------------------------------------------------------
Private Function BuildScreenBtn6() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn6()"

    On Error GoTo ErrorHandler

    Set BtnReport6 = New ClsUIButton
    
    With BtnReport6
        .Name = "BtnExpt6"
        MainFrame2.Buttons.Add BtnReport6
        .Height = BTN_EXP_6_HEIGHT
        .Left = BTN_EXP_6_LEFT
        .Top = BTN_EXP_6_TOP
        .Width = BTN_EXP_6_WIDTH
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Export 6" & vbCr & vbCr & "Export 6 description"
        .Visible = True
    End With
    
    
    BuildScreenBtn6 = True

Exit Function

ErrorExit:

    BuildScreenBtn6 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn7
' ---------------------------------------------------------------
Private Function BuildScreenBtn7() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn7()"

    On Error GoTo ErrorHandler

    Set BtnReport7 = New ClsUIButton
    
    With BtnReport7
        .Name = "BtnExpt7"
        MainFrame2.Buttons.Add BtnReport7
        .Height = BTN_EXP_7_HEIGHT
        .Left = BTN_EXP_7_LEFT
        .Top = BTN_EXP_7_TOP
        .Width = BTN_EXP_7_WIDTH
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Export 7" & vbCr & vbCr & "Export 7 description"
        .Visible = True
    End With
    
    
    BuildScreenBtn7 = True

Exit Function

ErrorExit:

    BuildScreenBtn7 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn8
' ---------------------------------------------------------------
Private Function BuildScreenBtn8() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn8()"

    On Error GoTo ErrorHandler

    Set BtnReport8 = New ClsUIButton
    
    With BtnReport8
        .Name = "BtnExpt8"
        MainFrame2.Buttons.Add BtnReport8
        .Height = BTN_EXP_8_HEIGHT
        .Left = BTN_EXP_8_LEFT
        .Top = BTN_EXP_8_TOP
        .Width = BTN_EXP_8_WIDTH
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Export 8" & vbCr & vbCr & "Export 8 description"
        .Visible = True
    End With
    
    
    BuildScreenBtn8 = True

Exit Function

ErrorExit:

    BuildScreenBtn8 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Method GetReportData
' Gets data for workflow list
'---------------------------------------------------------------
Public Function GetReportData(ReportNo As String) As Recordset
    Dim RstReport As Recordset
    Dim SQL As String
    
    Select Case ReportNo
        Case 1
            SQL = "Select " _
                & "   TblProject.ProjectName As Project, " _
                & "   TblCBSUser.UserName, " _
                & "   TblLender.[Name] As Lender, " _
                & "   TblWorkflow.Debt As Loan, " _
                & "   TblWorkflow.CBSComm As [Commission (%)], " _
                & "   TblWorkflow.ExitFee As [Exit Fee (%)], " _
                & "   TblWorkflow.CBSComm * TblWorkflow.Debt As Commission, " _
                & "   TblWorkflow.ExitFee * TblWorkflow.Debt As [Exit Fee] " _
                & " From " _
                & "   ((TblWorkflow Left Outer Join " _
                & "   TblProject On TblProject.ProjectNo = TblWorkflow.ProjectNo) Left Outer Join " _
                & "   TblLender On TblLender.LenderNo = TblWorkflow.LenderNo) Left Outer Join " _
                & "   TblCBSUser On TblCBSUser.CBSUserNo = TblProject.CaseManager " _
                & " Where " _
                & "   TblWorkflow.ProjectNo <> 0 " _
                & " Order By " _
                & "   TblWorkflow.ProjectNo"
            
        Case 2
            SQL = "Select " _
                    & "  Year(TblProject.CompleteDate) As [Year], " _
                    & "  Month(TblProject.CompleteDate) As [Month], " _
                    & "  Avg(TblWorkflow.CBSComm) As [Ave Commission] " _
                & "From " _
                    & "  TblProject Inner Join " _
                    & "  TblWorkflow On TblWorkflow.ProjectNo = TblProject.ProjectNo " _
                & "Where " _
                    & "  TblWorkflow.WorkflowType = 'enlender' " _
                    & "Group By " _
                    & "  Year(TblProject.CompleteDate), Month(TblProject.CompleteDate), " _
                    & "  TblProject.CompleteDate, TblWorkflow.WorkflowType "
                
        Case 3
            SQL = "Select " _
                    & "  TblCBSUser.UserName As [Case Manager], " _
                    & "  Avg(TblProject.CompleteDate - TblProject.StartDate) As [Average Duration] " _
                & "From " _
                    & "  TblProject Inner Join " _
                    & "  TblCBSUser On TblProject.CaseManager = TblCBSUser.CBSUserNo " _
                & "Where " _
                    & "  TblCBSUser.[Position] = 'Case Manager' And " _
                    & "  TblProject.CompleteDate Is Not Null " _
                    & "Group By " _
                    & "  TblCBSUser.UserName "
            
        Case 4
            SQL = "Select " _
                    & "  TblClient.[Name] As Client, " _
                    & "  Sum(TblWorkflow.Debt) As [Total Debt] " _
                & "From " _
                    & "  (TblProject Inner Join " _
                    & "  TblWorkflow On TblWorkflow.ProjectNo = TblProject.ProjectNo) Inner Join " _
                    & "  TblClient On TblClient.ClientNo = TblProject.ClientNo " _
                    & "Group By " _
                    & "  TblClient.[Name] "
            
    End Select
    
    Set RstReport = ModDatabase.SQLQuery(SQL)
    
    Set GetReportData = RstReport
    
End Function

