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
    
    Const StrPROCEDURE As String = "BtnProjectNewWFClick()"

    On Error GoTo ErrorHandler

    Set ActiveClient = New ClsClient
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
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Client"
        .Instructions = "Start typing the name of the Client and then select from the list. " _
                        & "Select 'New' to add a new Client"
        .Data = ModDatabase.SQLQuery("SELECT Name from TblClient")
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
        .SPV = ActiveSPV
        .DBSave
    End With
    
    With ActiveProject.ProjectWorkflow
        .Name = "Project"
        .WorkflowType = enProject
        .ActiveStep.Start
        .DBSave
        .DisplayForm
    
    End With

    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
    
GracefullExit:
    
    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    Set InputBox = Nothing
    
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
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Lender"
        .Instructions = "Select the Lender from the list.  Select New if the Lender you require is not listed"
        .Data = ModDatabase.SQLQuery("SELECT Name from TblLender")
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
    
    'get workflow type
    
    With ActiveLender
        .DBGet Picker.SelectedItem
        .DBSave
    End With
    
    Set ActiveWorkFlow = New ClsWorkflow
    
    With ActiveWorkFlow
        .Name = "Senior"
        .WorkflowType = enLender
        .Lender = ActiveLender
        .DBSave
    End With
    
    With ActiveProject
        .Workflows.Add ActiveWorkFlow
        .DBSave
    End With
    
    ActiveWorkFlow.DisplayForm
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
    
GracefullExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveLender = Nothing
    Set ActiveClient = Nothing
    
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
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnLenderOpenWFClick
' ---------------------------------------------------------------
Public Sub BtnLenderOpenWFClick(ByVal ScreenPage As enScreenPage, ByVal Index As String)
    If Not ModUIProjects.OpenLenderWF(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnProjectsClick
' ---------------------------------------------------------------
Public Sub BtnProjectsClick(ScreenPage As enScreenPage)
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
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
' BtnAdminUsersClick
' ---------------------------------------------------------------
Public Sub BtnAdminUsersClick(Optional ByVal ScreenPage As enScreenPage, Optional ByVal Index As String)
    CurrentUser.DisplayForm
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

