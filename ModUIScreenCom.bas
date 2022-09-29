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
