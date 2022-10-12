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
Public Sub ProcessBtnClicks(ButtonNo As String)
    Dim ErrNo As Integer
    Dim AryBtn() As String
    Dim Picker As ClsFrmPicker
    Dim BtnNo As EnumBtnNo
    Dim BtnIndex As Integer

    Const StrPROCEDURE As String = "ProcessBtnClicks()"
    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    AryBtn = Split(ButtonNo, ":")
    BtnNo = CInt(AryBtn(0))
    
    If UBound(AryBtn) = 1 Then BtnIndex = AryBtn(1)
    
    Select Case BtnNo
    
        Case enBtnForAction
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
        
        Case enBtnProjectsActive
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen(enActivePage) Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
            
        Case enBtnProjectsClosed
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen("Closed") Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMClients
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUICRM.BuildScreen(enCRMClient) Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMSPVs
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMContacts
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enCRMProjects
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
        
        Case enCRMLenders
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enDashboard
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enReports
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminUsers
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminEmailTs
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminDocuments
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminWorkflows
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminWFTypes
        
            ShtMain.Unprotect PROTECT_KEY
            ShtMain.[Button] = enBtnProjectsClosed

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen("Closed") Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminLists
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enAdminRoles
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR

            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnNewProjectWF
        
            ShtMain.Unprotect PROTECT_KEY

            Set ActiveClient = New ClsClient
            Set ActiveSPV = New ClsSPV
            Set ActiveWorkFlow = New ClsWorkflow
            Set ActiveProject = New ClsProject
            Set ActiveUser = New ClsCBSUser
            
            Set Picker = New ClsFrmPicker
            With Picker
                .Title = "Select Client"
                .Instructions = "Start typing the name of the Client and then select from the list. " _
                                & "Select 'New' to add a new Client"
                .Data = ModDatabase.SQLQuery("SELECT Name from TblClient")
                .ClearForm
                .Show = True
                If .CreateNew Then
                    ActiveUser.DBNew
                    .SelectedItem = ActiveUser.UserName
                End If
            
            End With
            
            ActiveClient.DBGet Picker.SelectedItem
            
            Set Picker = New ClsFrmPicker
            With Picker
                .Title = "Select SPV"
                .Instructions = "Start typing the name of the SPV and then select from the list. " _
                                & "Select 'New' to add a new SPV"
                .Data = ModDatabase.SQLQuery("SELECT Name from TblSPV")
                .ClearForm
                .Show = True
                If .CreateNew Then
                    ActiveSPV.DBNew
                    .SelectedItem = ActiveSPV.Name
                End If
            End With
            
            ActiveSPV.DBGet Picker.SelectedItem
            ActiveClient.SPVs.Add ActiveSPV
            
            Set Picker = New ClsFrmPicker
            With Picker
                .Title = "Select Case Manager"
                .Instructions = "Start typing the name of the Case Manager and then select from the list. " _
                                & "Select 'New' to add a new User"
                .Data = ModDatabase.SQLQuery("SELECT Username from TblCBSUser")
                .ClearForm
                .Show = True
                If .CreateNew Then
                    ActiveUser.DBNew
                    .SelectedItem = ActiveSPV.Name
                End If
            End With
            
            ActiveUser.DBGet Picker.SelectedItem
            
            With ActiveProject
                .ProjectWorkflow.Name = "Project"
                .CaseManager = ActiveUser
                .DBSave
            End With
            
            ActiveSPV.Projects.Add ActiveProject
            
            Debug.Assert ActiveClient.SPVs.Count > 0
            
            ActiveClient.DBSave
            
            With ActiveWorkFlow
                .Name = "Project"
                .WorkflowType = enProject
                .DBSave
            End With
            
                ActiveProject.ProjectWorkflow = ActiveWorkFlow
                Debug.Assert ActiveProject.ProjectWorkflow.Steps.Count > 0
                        
                FrmProject.ShowForm
            
            Debug.Assert Not ActiveClient Is Nothing
            Debug.Assert ActiveWorkFlow.Steps.Count > 0
            
            
            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen(enActivePage) Then Err.Raise HANDLED_ERROR
            
            ShtMain.Protect PROTECT_KEY
            
            Set ActiveWorkFlow = Nothing
            Set ActiveSPV = Nothing
            Set ActiveProject = Nothing
            Set Picker = Nothing
            Set ActiveClient = Nothing
            Set ActiveUser = Nothing
            
            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnNewLenderWF
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            ShtMain.Protect PROTECT_KEY
            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnExit
        
            ShtMain.Unprotect PROTECT_KEY

            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            ShtMain.Protect PROTECT_KEY
            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

        Case enBtnOpenProject

            ShtMain.Unprotect PROTECT_KEY
            Set ActiveProject = New ClsProject
            Set ActiveSPV = New ClsSPV
            Set ActiveClient = New ClsClient
            
            ActiveProject.DBGet BtnIndex
            Set ActiveSPV = ActiveProject.Parent
            Set ActiveClient = ActiveProject.Parent.Parent
            
            FrmProject.ShowForm
            
            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUIProjects.BuildScreen(enActivePage) Then Err.Raise HANDLED_ERROR
            
            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
            
            Set ActiveProject = Nothing
            Set ActiveSPV = Nothing
            Set ActiveClient = Nothing

        Case enBtnOpenCRM
        
            ShtMain.Unprotect PROTECT_KEY

            Set ActiveClient = New ClsClient
            
            With ActiveClient
                .DBGet BtnIndex
                .DisplayForm
            End With
            
            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUICRM.BuildScreen(enCRMClient) Then Err.Raise HANDLED_ERROR
            
            ShtMain.Protect PROTECT_KEY

            If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY

            Set ActiveClient = Nothing

        Case enBtnNewClient
        
            ShtMain.Unprotect PROTECT_KEY

            Set ActiveClient = New ClsClient
            With ActiveClient
                .DBNew
            End With
            
            If Not ResetScreen Then Err.Raise HANDLED_ERROR
            If Not ModUICRM.BuildScreen(enCRMClient) Then Err.Raise HANDLED_ERROR

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
