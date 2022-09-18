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
' New Project Workflow Button
' ---------------------------------------------------------------
Public Function BuildScreenBtn1() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn1()"

    On Error GoTo ErrorHandler

    Set BtnNewProject = New ClsUIMenuItem

    With BtnNewProject

        .Height = BTN_MAIN_1_HEIGHT
        .Left = BTN_MAIN_1_LEFT
        .Top = BTN_MAIN_1_TOP
        .Width = BTN_MAIN_1_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnNewProject & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Project Workflow"
    End With

    MainFrame.Menu.AddItem BtnNewProject

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
' New Project Workflow Button
' ---------------------------------------------------------------
Public Function BuildScreenBtn2() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn2()"

    On Error GoTo ErrorHandler

    Set BtnNewLender = New ClsUIMenuItem

    With BtnNewLender

        .Height = BTN_MAIN_2_HEIGHT
        .Left = BTN_MAIN_2_LEFT
        .Top = BTN_MAIN_2_TOP
        .Width = BTN_MAIN_2_WIDTH
        .Name = "BtnMain2"
        .OnAction = "'ModUIScreenCom.ProcessBtnPress(" & enBtnNewLender & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Lender Workflow"
    End With

    MainFrame.Menu.AddItem BtnNewLender

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
' ProcessBtnPress
' Receives all button presses and processes
' ---------------------------------------------------------------
Public Sub ProcessBtnPress(ButtonNo As Integer)
    Dim ErrNo As Integer
    Dim Response As Integer
    Dim NewWorkFlow As ClsWorkflow
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
        
        Case enBtnNewProject
                
            FrmWorkflow.Show
            
        Case BtnNewLender
        
            FrmWorkflow.Show
                        
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

