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
    
    MainFrame.ReOrder
'    MainFrame2.ReOrder
           
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
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame1() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame1()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = MAIN_FRAME_HEIGHT
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
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame2() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame2()"

    On Error GoTo ErrorHandler

    Set MainFrame2 = New ClsUIFrame
    MainScreen.Frames.AddItem MainFrame2, "Main Frame 2"
    
    'add main frame
    With MainFrame2
        .Top = MAIN_FRAME_2_TOP
        .Left = MAIN_FRAME_2_LEFT
        .Width = MAIN_FRAME_2_WIDTH
        .Height = MAIN_FRAME_2_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        
        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame 2 Header"
            .Text = "Data Exports"
            .Style = HEADER_STYLE
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
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildScreenBtn1() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn1()"

    On Error GoTo ErrorHandler

    Set BtnReport1 = New ClsUIMenuItem

    
    With BtnReport1
        .Name = "BtnRep1"
        MainFrame.Menu.AddItem BtnReport1
        .Height = BTN_REP_1_HEIGHT
        .Left = BTN_REP_1_LEFT
        .Top = BTN_REP_1_TOP
        .Width = BTN_REP_1_WIDTH
        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport1 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Expiry Report" & vbCr & vbCr & "Members whose qualifications are close to or have expired"
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
' Builds Report Button 2
' ---------------------------------------------------------------
Private Function BuildScreenBtn2() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn2()"

    On Error GoTo ErrorHandler

    Set BtnReport2 = New ClsUIMenuItem

    
    With BtnReport2
        .Name = "BtnRep2"
        MainFrame.Menu.AddItem BtnReport2
        .Height = BTN_REP_2_HEIGHT
        .Left = BTN_REP_2_LEFT
        .Top = BTN_REP_2_TOP
        .Width = BTN_REP_2_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport2 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Qualification Dates Report" & vbCr & vbCr & "Reports all certification dates for a selected Member "
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
' Builds Report Button 3
' ---------------------------------------------------------------
Private Function BuildScreenBtn3() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn3()"

    On Error GoTo ErrorHandler

    Set BtnReport3 = New ClsUIMenuItem

    
    With BtnReport3
        .Name = "BtnRep3"
        MainFrame2.Menu.AddItem BtnReport3
        .Height = BTN_REP_3_HEIGHT
        .Left = BTN_REP_3_LEFT
        .Top = BTN_REP_3_TOP
        .Width = BTN_REP_3_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport3 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Member Data" & vbCr & vbCr & "Exports all Member data to Excel"
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
' Builds Report Button 4
' ---------------------------------------------------------------
Private Function BuildScreenBtn4() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn4()"

    On Error GoTo ErrorHandler

    Set BtnReport4 = New ClsUIMenuItem

    
    With BtnReport4
        .Name = "BtnRep4"
        MainFrame2.Menu.AddItem BtnReport4
        .Height = BTN_REP_4_HEIGHT
        .Left = BTN_REP_4_LEFT
        .Top = BTN_REP_4_TOP
        .Width = BTN_REP_4_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport4 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Qualification Dates" & vbCr & vbCr & "Exports all qualification date data to Excel"
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
' Builds Certs matrix
' ---------------------------------------------------------------
Private Function BuildScreenBtn5() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn5()"

    On Error GoTo ErrorHandler

    Set BtnReport5 = New ClsUIMenuItem

    
    With BtnReport5
        .Name = "BtnRep5"
        MainFrame.Menu.AddItem BtnReport5
        .Height = BTN_REP_5_HEIGHT
        .Left = BTN_REP_5_LEFT
        .Top = BTN_REP_5_TOP
        .Width = BTN_REP_5_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport5 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Certification Matrix" & vbCr & vbCr & "Excel Matrix of Member's certifications and QIP status"
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
' Builds Certs matrix
' ---------------------------------------------------------------
Private Function BuildScreenBtn6() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn6()"

    On Error GoTo ErrorHandler

    Set BtnReport6 = New ClsUIMenuItem

    
    With BtnReport6
        .Name = "BtnRep6"
        MainFrame.Menu.AddItem BtnReport6
        .Height = BTN_REP_6_HEIGHT
        .Left = BTN_REP_6_LEFT
        .Top = BTN_REP_6_TOP
        .Width = BTN_REP_6_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport6 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "ERC Requirements" & vbCr & vbCr & "Training Certification status of roles for quarterly reporting"
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
' Not Enrolled report button
' ---------------------------------------------------------------
Private Function BuildScreenBtn7() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn7()"

    On Error GoTo ErrorHandler

    Set BtnReport7 = New ClsUIMenuItem

    
    With BtnReport7
        .Name = "BtnRep7"
        MainFrame.Menu.AddItem BtnReport7
        .Height = BTN_REP_7_HEIGHT
        .Left = BTN_REP_7_LEFT
        .Top = BTN_REP_7_TOP
        .Width = BTN_REP_7_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport7 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Not Enrolled Report" & vbCr & vbCr & "Returns candidates that are not enrolled on a course"
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
' Not Enrolled report button
' ---------------------------------------------------------------
Private Function BuildScreenBtn8() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn8()"

    On Error GoTo ErrorHandler

    Set BtnReport8 = New ClsUIMenuItem

    With BtnReport8
        .Name = "BtnRep8"
        MainFrame.Menu.AddItem BtnReport8
        .Height = BTN_REP_8_HEIGHT
        .Left = BTN_REP_8_LEFT
        .Top = BTN_REP_8_TOP
        .Width = BTN_REP_8_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport8 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Past CDC End Date" & vbCr & vbCr & "Returns candidates that have not completed CDC and are past End Date"
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
' BuildScreenBtn9
' Not Enrolled report button
' ---------------------------------------------------------------
Private Function BuildScreenBtn9() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn9()"

    On Error GoTo ErrorHandler

    Set BtnReport9 = New ClsUIMenuItem

    
    With BtnReport9
        .Name = "BtnRep9"
        MainFrame.Menu.AddItem BtnReport9
        .Height = BTN_REP_9_HEIGHT
        .Left = BTN_REP_9_LEFT
        .Top = BTN_REP_9_TOP
        .Width = BTN_REP_9_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport9 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Not Enrolled and Not QIP" & vbCr & vbCr & "Returns candidates that are not qualified in post and are not enrolled on a CDC"
    End With
    
    BuildScreenBtn9 = True

Exit Function

ErrorExit:

    BuildScreenBtn9 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildScreenBtn10
' Not Enrolled report button
' ---------------------------------------------------------------
Private Function BuildScreenBtn10() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn10()"

    On Error GoTo ErrorHandler

    Set BtnReport10 = New ClsUIMenuItem

    
    With BtnReport10
        .Name = "BtnRep10"
        MainFrame.Menu.AddItem BtnReport10
        .Height = BTN_REP_10_HEIGHT
        .Left = BTN_REP_10_LEFT
        .Top = BTN_REP_10_TOP
        .Width = BTN_REP_10_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport10 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Not Enrolled and QIP" & vbCr & vbCr & "Returns candidates that are qualified in post but not enrolled on a CDC"
    End With
    
    BuildScreenBtn10 = True

Exit Function

ErrorExit:

    BuildScreenBtn10 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildScreenBtn11
' Not Enrolled report button
' ---------------------------------------------------------------
Private Function BuildScreenBtn11() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn11()"

    On Error GoTo ErrorHandler

    Set BtnReport11 = New ClsUIMenuItem

    
    With BtnReport11
        .Name = "BtnRep11"
        MainFrame.Menu.AddItem BtnReport11
        .Height = BTN_REP_11_HEIGHT
        .Left = BTN_REP_11_LEFT
        .Top = BTN_REP_11_TOP
        .Width = BTN_REP_11_WIDTH
       .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnReport11 & ")'"
        .UnSelectStyle = TOOL_BUTTON
        .Selected = False
        .Text = "Quals Gained Between Two Dates" & vbCr & vbCr & "Returns the number of quals achieved on the FD between two given dates"
    End With
    
    BuildScreenBtn11 = True

Exit Function

ErrorExit:

    BuildScreenBtn11 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function











