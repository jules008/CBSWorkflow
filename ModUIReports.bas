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

    Set BtnReport1 = New ClsUIButton
    
    With BtnReport1
        .Name = "BtnRep1"
        MainFrame.Buttons.Add BtnReport1
        .Height = BTN_REP_1_HEIGHT
        .Left = BTN_REP_1_LEFT
        .Top = BTN_REP_1_TOP
        .Width = BTN_REP_1_WIDTH
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
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
' Adds the button to switch order list between open and closed orders
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
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Expiry Report" & vbCr & vbCr & "Members whose qualifications are close to or have expired"
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
' Adds the button to switch order list between open and closed orders
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
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Expiry Report" & vbCr & vbCr & "Members whose qualifications are close to or have expired"
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
' Adds the button to switch order list between open and closed orders
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
        .OnAction = ""
        .UnSelectStyle = BTN_MAIN_STYLE
        .Selected = False
        .Text = "Expiry Report" & vbCr & vbCr & "Members whose qualifications are close to or have expired"
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

