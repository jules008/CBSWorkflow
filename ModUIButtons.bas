Attribute VB_Name = "ModUIButtons"
'===============================================================
' Module ModUIButtons
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 02 Oct 22
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIButtons"

' ===============================================================
' ProcessBtnClicks
' Processes all button presses in application
' ---------------------------------------------------------------
Public Sub ProcessBtnClicks(ButtonIndex As EnumBtnNo)
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "ProcessBtnClicks()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    Select Case ButtonIndex
    
        Case enBtnForAction
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
        
        Case enBtnProjectsActive
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen("Active") Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
            
        Case enBtnProjectsClosed
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen("Closed") Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMClients
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMSPVs
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMContacts
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMProjects
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
        
        Case enCRMLenders
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enDashboard
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enReports
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminUsers
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminEmailTs
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminDocuments
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminWorkflows
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminWFTypes
        
            ShtMain.Unprotect PROTECT_KEY
            ShtMain.[Button] = enBtnProjectsClosed

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen("Closed") Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminLists
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminRoles
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnNewProjectWF
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnNewLenderWF
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnExit
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Unprotect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

    End Select


GracefulExit:


Exit Sub

ErrorExit:

    '***CleanUpCode***

Exit Sub

ErrorHandler:
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
