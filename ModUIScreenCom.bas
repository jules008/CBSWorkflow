Attribute VB_Name = "ModUIScreenCom"
'===============================================================
' Module ModUIScreenCom
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Nov 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIScreenCom"

' ===============================================================
' BuildScreenBtn1
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Public Function BuildScreenBtn1() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn1()"

    On Error GoTo ErrorHandler

    Set BtnNewWorkflow = New ClsUIButton

    With BtnNewWorkflow

        .Height = BTN_MAIN_1_HEIGHT
        .Left = BTN_MAIN_1_LEFT
        .Top = BTN_MAIN_1_TOP
        .Width = BTN_MAIN_1_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnNewWorkflow & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Workflow"
    End With

    MainFrame.Menu.AddButton BtnNewWorkflow

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

'' ===============================================================
'' BuildScreenBtn2
'' Manage dates button
'' ---------------------------------------------------------------
'Public Function BuildScreenBtn2() As Boolean
'
'    Const StrPROCEDURE As String = "BuildScreenBtn2()"
'
'    On Error GoTo ErrorHandler
'
'    Set BtnCopyEmail = New ClsUIButton
'
'    With BtnCopyEmail
'
'        .Height = BTN_MAIN_2_HEIGHT
'        .Left = BTN_MAIN_2_LEFT
'        .Top = BTN_MAIN_2_TOP
'        .Width = BTN_MAIN_2_WIDTH
'        .Name = "BtnMain2"
'        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnCopyEmail & ")'"
'
'        .UnSelectStyle = TOOL_BUTTON
'        .Selected = False
'        .Text = "Copy Email"
'    End With
'
'    MainFrame.Menu.AddButton BtnCopyEmail
'
'    BuildScreenBtn2 = True
'
'Exit Function
'
'ErrorExit:
'
'    BuildScreenBtn2 = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
'
'' ===============================================================
'' BuildScreenBtn3
'' Manage dates button
'' ---------------------------------------------------------------
'Public Function BuildScreenBtn3() As Boolean
'
'    Const StrPROCEDURE As String = "BuildScreenBtn3()"
'
'    On Error GoTo ErrorHandler
'
'    Set BtnViewDates = New ClsUIButton
'
'    With BtnViewDates
'
'        .Height = BTN_MAIN_3_HEIGHT
'        .Left = BTN_MAIN_3_LEFT
'        .Top = BTN_MAIN_3_TOP
'        .Width = BTN_MAIN_3_WIDTH
'        .Name = "BtnMain3"
'        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnDates & ")'"
'
'        .UnSelectStyle = GENERIC_BUTTON
'        .Selected = False
'        .Text = "View Dates"
'    End With
'
'    MainFrame.Menu.AddButton BtnViewDates
'
'    BuildScreenBtn3 = True
'
'Exit Function
'
'ErrorExit:
'
'    BuildScreenBtn3 = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
'' ===============================================================
'' BuildScreenBtn4
'' Manage dates button
'' ---------------------------------------------------------------
'Public Function BuildScreenBtn4() As Boolean
'
'    Const StrPROCEDURE As String = "BuildScreenBtn4()"
'
'    On Error GoTo ErrorHandler
'
'    Set BtnMbrSummary = New ClsUIButton
'
'    With BtnMbrSummary
'
'        .Height = BTN_MAIN_4_HEIGHT
'        .Left = BTN_MAIN_4_LEFT
'        .Top = BTN_MAIN_4_TOP
'        .Width = BTN_MAIN_4_WIDTH
'        .Name = "BtnMain4"
'        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnSummary & ")'"
'
'        .UnSelectStyle = GENERIC_BUTTON
'        .Selected = False
'        .Text = "CDC" & vbCr & "Summary"
'    End With
'
'    MainFrame.Menu.AddButton BtnMbrSummary
'
'    BuildScreenBtn4 = True
'
'Exit Function
'
'ErrorExit:
'
'    BuildScreenBtn4 = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
'' ===============================================================
'' BuildScreenBtn5
'' Manage dates button
'' ---------------------------------------------------------------
'Public Function BuildScreenBtn5() As Boolean
'
'    Const StrPROCEDURE As String = "BuildScreenBtn5()"
'
'    On Error GoTo ErrorHandler
'
'    Set BtnCDCLookUp = New ClsUIButton
'
'    With BtnCDCLookUp
'
'        .Height = BTN_MAIN_5_HEIGHT
'        .Left = BTN_MAIN_5_LEFT
'        .Top = BTN_MAIN_5_TOP
'        .Width = BTN_MAIN_5_WIDTH
'        .Name = "BtnMain5"
'        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnCDCLookUp & ")'"
'
'        .UnSelectStyle = TOOL_BUTTON
'        .Selected = False
'        .Text = "CDC# Look Up "
'    End With
'
'    MainFrame.Menu.AddButton BtnCDCLookUp
'
'    BuildScreenBtn5 = True
'
'Exit Function
'
'ErrorExit:
'
'    BuildScreenBtn5 = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
'' ===============================================================
'' DestroyMainScreen
'' Destroys the main screen objects
'' ---------------------------------------------------------------
'Public Function DestroyMainScreen() As Boolean
'    Dim Frame As ClsUIFrame
'
'    Const StrPROCEDURE As String = "DestroyMainScreen()"
'
'    On Error GoTo ErrorHandler
'
'    Set Frame = New ClsUIFrame
'
'    For Each Frame In MainScreen.Frames
'        If Frame.Name <> "MenuBar" Then
'            MainScreen.Frames.RemoveItem Frame.Name
'        End If
'    Next
'
'    If Not MainFrame Is Nothing Then MainFrame.Visible = False
'
'    Set MainFrame = Nothing
'
'    Set Frame = Nothing
'
'    DestroyMainScreen = True
'
'Exit Function
'
'ErrorExit:
'
'    Set Frame = Nothing
'
'    DestroyMainScreen = False
'
'Exit Function
'
'ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
' ===============================================================
' ProcessBtnPress
' Receives all button presses and processes
' ---------------------------------------------------------------
Public Sub ProcessBtnPress(ButtonNo As Integer)
    Dim ErrNo As Integer
    Dim Response As Integer
    Dim NewWorkFlow As ClsWorkflow
'    Dim SelMember As ClsMember
    Dim SelWorkflow As String
    Dim AllWorkflows As ClsWorkflows

    Const StrPROCEDURE As String = "ProcessBtnPress()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

        Application.StatusBar = ""
                
'        Set AllWorkflows = New ClsWorkflows
        Set ActiveWorkFlow = Nothing
        Set ActiveWorkFlow = New ClsWorkflow
        
        Select Case ButtonNo
        
        Case enBtnNewWorkflow
                
            With ActiveWorkFlow
                .Name = "initial"
                .DBSave
                
                FrmWorkflow.ShowForm
                If Not ModUIActive.RefreshList Then Err.Raise HANDLED_ERROR
                .DBSave
            End With
                        
    End Select

GracefulExit:

'    Set SelMember = Nothing
    Set NewWorkFlow = Nothing
'    Workflows.Terminate
'    Set Workflows = Nothing
    
    Application.DisplayAlerts = True

Exit Sub

ErrorExit:
'    Set SelMember = Nothing
    Set NewWorkFlow = Nothing
'    Workflows.Terminate
'    Set Workflows = Nothing
    
    Application.DisplayAlerts = True

Exit Sub

ErrorHandler:
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler Err.Number, Err.Description
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'' ===============================================================
'' SortBy
'' Sorts cols by selected field
'' ---------------------------------------------------------------
'Private Function SortBy(StrCol As String, Optional SelScreen As String) As Boolean
'    Dim StrSort As String
'
'    Const StrPROCEDURE As String = "SortBy()"
'
'    On Error GoTo ErrorHandler
'
'    Select Case StrCol
'        Case Is = "Name"
'            StrSort = "TblMember.DisplayName"
'        Case Is = "SSN"
'            StrSort = "TblMember.SSN"
'        Case Is = "Student ID"
'            StrSort = "TblMember.StudentID"
'        Case Is = "Watch"
'            StrSort = "TblMember.Watch"
'        Case Is = "DoD Certification"
'            StrSort = "TblDoDCert.CertName"
'        Case Is = "Step No"
'            StrSort = "TblWorkflow.CurrentStep"
'        Case Is = "Step Name"
'            StrSort = "TblStep.StepName"
'        Case Is = "Status"
'            StrSort = "TblWorkflow.Status"
'    End Select
'
'    Select Case SelScreen
'        Case "UIActive"
'            ModUIActive.RefreshList StrSort
'        Case "UIForAction"
'            ModUIForAction.RefreshList StrSort
'
'    End Select
'
'    SortBy = True
'
'
'Exit Function
'
'ErrorExit:
'
'    '***CleanUpCode***
'    SortBy = False
'
'Exit Function
'
'ErrorHandler:
'    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'
'
