Attribute VB_Name = "ModUIButtonScripts"
'===============================================================
' Module ModUIButtonScripts
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Oct 22
'===============================================================

 Option Explicit

Private Const StrMODULE As String = "ModUIButtonScripts"

' ===============================================================
' BtnProjectNewWFClick
' Generates new project workflow
' ---------------------------------------------------------------
Public Function BtnProjectNewWFClick(ScreenPage As enScreenPage) As Boolean
    Dim Picker As ClsFrmPicker
    Dim InputBox As ClsInputBox
    Dim ProjectName As String
    Dim RstSource As Recordset
    
    Const StrPROCEDURE As String = "BtnProjectNewWFClick()"

    On Error GoTo ErrorHandler

    Set ActiveClient = New ClsClient
    Set ActiveContact = New ClsContact
    Set ActiveSPV = New ClsSPV
    Set ActiveWorkFlow = New ClsWorkflow
    Set ActiveProject = New ClsProject
    Set ActiveUser = New ClsCBSUser
    Set InputBox = New ClsInputBox
    
    'get project name
    With InputBox
        .Title = "Enter Project Name"
        .Instructions = "Enter a meaningful name for the project"
        .ClearForm
        .Show
        ProjectName = .ReturnValue
    End With
    
    If InputBox.ReturnValue = "" Then
        MsgBox "A Project Name is needed to continue, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    'Get Client
    If CurrentUser.UserLvl = enCaseMgr Then
        Set RstSource = ModDatabase.SQLQuery("Select " _
                                        & "    TblClient.Name " _
                                        & "From " _
                                        & "    TblAccessControl Right Join " _
                                        & "    TblClient On TblClient.ClientNo = TblAccessControl.EntityNo " _
                                        & "Where " _
                                        & "    TblAccessControl.Entity = 'Client' And " _
                                        & "    TblAccessControl.UserNo = " _
                                        & CurrentUser.CBSUserNo)
    Else
        Set RstSource = ModDatabase.SQLQuery("SELECT Name from TblClient")
    End If
    
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Client"
        .Instructions = "Start typing the name of the Client and then select from the list. " _
                        & "Select 'New' to add a new Client"
        .Data = RstSource
        .ClearForm
        .Show = True
        If .CreateNew Then
            ActiveClient.DBNew
            .SelectedItem = ActiveClient.Name
        End If
    
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    With ActiveClient
        .DBGet Picker.SelectedItem
        .DBSave
    End With
    
    'Get Client Primary contact
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Client Primary Contact"
        .Instructions = "Start typing the name of the Client Contact and then select from the list. " _
                        & "Select 'New' to add a new Client Contact"
        .Data = ModDatabase.SQLQuery("SELECT ContactName from TblContact")
        .ClearForm
        .Show = True
        If .CreateNew Then
            ActiveClient.Contacts.PrimaryContact.DBNew "Client", ActiveClient.Name
            .SelectedItem = ActiveClient.Contacts.PrimaryContact.ContactName
        End If
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    With ActiveContact
        .DBGet Picker.SelectedItem
        .Organisation = ActiveClient.Name
        .DBSave
    End With
    
    With ActiveClient.Contacts
        .Add ActiveContact
        .PrimaryContact = ActiveContact
    End With
    
    ActiveClient.DBSave
    
    'Get SPV
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
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
   
    With ActiveSPV
        .DBGet Picker.SelectedItem
        .DBSave
    End With
    
    'get Case Manager
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
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    With ActiveUser
        .DBGet Picker.SelectedItem
        .DBSave
    End With
    
    ActiveClient.SPVs.Add ActiveSPV
    
    With ActiveProject
        .ProjectWorkflow.Name = "Project"
        .CaseManager = ActiveUser
        .ProjectName = ProjectName
        .Client = ActiveClient
        .StartDate = Now
        .SPV = ActiveSPV
        .ConsComenceDte = Format(Now, "dd mmm yy")
        .CBSCommission = 0
        .CBSCommPC = 0
        .Debt = 0
        .ExitFee = 0
        .ExitFeePC = 0
        .DBSave
    End With
    
    With ActiveProject.ProjectWorkflow
        .Name = "Project"
        .WorkflowType = enProject
        .DBSave
        .ActiveStep.Start
        .DisplayProjectForm
    
    End With

    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage, False) Then Err.Raise HANDLED_ERROR
    
GracefullExit:
    
    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    Set InputBox = Nothing
    Set ActiveContact = Nothing
    Set RstSource = Nothing
    
    BtnProjectNewWFClick = True

Exit Function

ErrorExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    Set InputBox = Nothing
    Set RstSource = Nothing
    
    BtnProjectNewWFClick = False

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
' BtnLenderNewWFClick
' Generates new lender workflow
' ---------------------------------------------------------------
Public Function BtnLenderNewWFClick(ScreenPage As enScreenPage) As Boolean
    Dim Picker As ClsFrmPicker
    Dim SQL As String
    Dim RstSource As Recordset
    
    Const StrPROCEDURE As String = "BtnLenderNewWFClick()"

    On Error GoTo ErrorHandler
    
    Set ActiveClient = New ClsClient
    Set ActiveProject = New ClsProject
    Set ActiveSPV = New ClsSPV
    Set ActiveWorkFlow = New ClsWorkflow
    Set ActiveProject = New ClsProject
    Set ActiveLender = New ClsLender
    
    'Get Project
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Project"
        .Instructions = "Select the Project from the list that you would like to add a Lender Workflow to"
        .Data = ModDatabase.SQLQuery("SELECT ProjectName from TblProject")
        .ClearForm
        .Show = True
        If .CreateNew Then
            ActiveProject.DBNew
            .SelectedItem = ActiveProject.ProjectName
        End If
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    With ActiveProject
        .DBGet Picker.SelectedItem
        .DBSave
    End With

    'get lender
    If CurrentUser.UserLvl = enCaseMgr Then
        Set RstSource = ModDatabase.SQLQuery("Select " _
                                        & "    TblLender.Name " _
                                        & "From " _
                                        & "    TblAccessControl Right Join " _
                                        & "    TblLender On TblLender.LenderNo = TblAccessControl.EntityNo " _
                                        & "Where " _
                                        & "    TblAccessControl.Entity = 'Lender' And " _
                                        & "    TblAccessControl.UserNo = " _
                                        & CurrentUser.CBSUserNo)
    Else
        Set RstSource = ModDatabase.SQLQuery("SELECT Name from TblLender")
    End If
    
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Lender"
        .Instructions = "Select the Lender from the list.  Select New if the Lender you require is not listed"
        .Data = RstSource
        .ClearForm
        .Show = True
        If .CreateNew Then
            ActiveLender.DBNew
            .SelectedItem = ActiveLender.Name
        End If
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefullExit
    End If
    
    With ActiveLender
        .DBGet Picker.SelectedItem
        .DBSave
    End With
    
    Set ActiveWorkFlow = New ClsWorkflow
    
    With ActiveWorkFlow
        .DisplayWFSelectForm
        If .Name = "" Then
            MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
            GoTo GracefullExit
        End If
        .WorkflowType = enLender
        .Lender = ActiveLender
        .DBSave
    End With
    
    With ActiveProject
        .Workflows.Add ActiveWorkFlow
        .DBSave
    End With
    
    ActiveWorkFlow.DisplayLenderForm
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage, False) Then Err.Raise HANDLED_ERROR
    
GracefullExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveLender = Nothing
    Set ActiveClient = Nothing
    Set RstSource = Nothing
    
    BtnLenderNewWFClick = True

Exit Function

ErrorExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveLender = Nothing
    Set ActiveClient = Nothing
    Set RstSource = Nothing
    
    BtnLenderNewWFClick = False

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
' BtnProjectOpenWFClick
' ---------------------------------------------------------------
Public Sub BtnProjectOpenWFClick(ByVal ScreenPage As enScreenPage, ByVal Index As String)
    If Not ModUIProjects.OpenProjectWF(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage, False) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnLenderOpenWFClick
' ---------------------------------------------------------------
Public Sub BtnLenderOpenWFClick(ByVal ScreenPage As enScreenPage, ByVal Index As String)
    If Not ModUIProjects.OpenLenderWF(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.RefreshList(ScreenPage, True) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnProjectsClick
' ---------------------------------------------------------------
Public Sub BtnProjectsClick(ScreenPage As enScreenPage)
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage, False) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnCRMClick
' ---------------------------------------------------------------
Public Sub BtnCRMClick(ScreenPage As enScreenPage)
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUICRM.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnCRMOpenItem
' ---------------------------------------------------------------
Public Sub BtnCRMOpenItem(ByVal ScreenPage As enScreenPage, Optional ByVal Index As String)
    If Not ModUICRM.OpenItem(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUICRM.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnAdminOpenItem
' ---------------------------------------------------------------
Public Sub BtnAdminOpenItem(ByVal ScreenPage As enScreenPage, Optional ByVal Index As String)
    If Not ModUIAdmin.OpenItem(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIAdmin.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnCRMContCalImport
' ---------------------------------------------------------------
Public Sub BtnCRMContCalImport(ByVal ScreenPage As enScreenPage, Optional ByVal Index As String)
    If Not ModUICRM.CalendlyImport() Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUICRM.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnCRMContShwLeads
' ---------------------------------------------------------------
Public Sub BtnCRMContShwLeads(ByVal ScreenPage As enScreenPage, Optional ByVal Index As String)
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUICRM.BuildScreen(ScreenPage, "ContactType:Lead") Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnDashboardClick
' ---------------------------------------------------------------
Public Sub BtnDashboardClick()
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIDashboard.BuildScreen() Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnReportsClick
' ---------------------------------------------------------------
Public Sub BtnReportsClick()
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIReports.BuildScreen() Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnAdminClick
' ---------------------------------------------------------------
Public Sub BtnAdminClick(ByVal ScreenPage As enScreenPage)
    Dim Picker As ClsFrmPicker
    Dim StrFilter As String
    Dim RstFilter As Recordset
    
    If ScreenPage = enScrAdminWorkflows Then
        Set Picker = New ClsFrmPicker
        With Picker
            .Title = "Select workflow script"
            .Instructions = "Select the workflow script you would like to view."
            .Data = ModDatabase.SQLQuery("SELECT SecondTier from TblWorkflowTable")
            .ClearForm
            .Show = True
        End With
        
        If Picker.SelectedItem = "" Then
            MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
            GoTo GracefullExit
        End If
        
        Set RstFilter = ModDatabase.SQLQuery("SELECT WFNo FROM TblWorkflowTable WHERE SecondTier = '" & Picker.SelectedItem & "'")
        StrFilter = "WorkflowNo:" & RstFilter!WFNo
    End If
    
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIAdmin.BuildScreen(ScreenPage, StrFilter) Then Err.Raise HANDLED_ERROR
    
    Set RstFilter = Nothing
GracefullExit:

    Set Picker = Nothing
End Sub

' ===============================================================
' BtnCommsToDoClick
' ---------------------------------------------------------------
Public Sub BtnCommsToDoClick()
    Dim Contact As New ClsContact
    
    Contact.DisplayCommForm
    ButtonFrame.Buttons("BtnMain3").Badge.Text = Contact.CommsNo
    Set Contact = Nothing
End Sub

' ===============================================================
' BtnReportSel
' Selects a report from report screen
' ---------------------------------------------------------------
Public Function BtnReportSel(ByRef ReportNo As String) As Boolean
    Dim RstReportData As Recordset
    
    Const StrPROCEDURE As String = "BtnReportSel()"

    On Error GoTo ErrorHandler
    
    Set RstReportData = ModUIReports.GetReportData(ReportNo)
    
    Select Case ReportNo
        Case "1"
            If Not ModReport.Report1(RstReportData) Then Err.Raise HANDLED_ERROR
    
        Case "2"
            If Not ModReport.Report2(RstReportData) Then Err.Raise HANDLED_ERROR
    
        Case "3"
            If Not ModReport.Report3(RstReportData) Then Err.Raise HANDLED_ERROR
    
        Case "4"
            If Not ModReport.Report4(RstReportData) Then Err.Raise HANDLED_ERROR
    
    End Select
    
    BtnReportSel = True
    Set RstReportData = Nothing
    
Exit Function
    
ErrorExit:
    
    Set RstReportData = Nothing
    BtnReportSel = False

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
' BtnExitClick
' ---------------------------------------------------------------
Public Sub BtnExitClick()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME)

    If Response = 6 Then

        If Workbooks.Count = 1 Then
            With Application
                .DisplayAlerts = False
                .Quit
                .DisplayAlerts = True
            End With
        Else
            ThisWorkbook.Close savechanges:=False
        End If

    End If
End Sub

