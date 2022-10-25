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
    
    Const StrPROCEDURE As String = "BtnProjectNewWFClick()"

    On Error GoTo ErrorHandler

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
            ActiveClient.DBNew
            .SelectedItem = ActiveClient.Name
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
                
    ActiveWorkFlow.DBNew
    
    Debug.Assert Not ActiveClient Is Nothing
    Debug.Assert ActiveWorkFlow.Steps.Count > 0
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    
    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR

    BtnProjectNewWFClick = True

Exit Function

ErrorExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    
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
    
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select Project"
        .Instructions = "Select the Project from the list that you would like to add a Lender Workflow to"
        .Data = ModDatabase.SQLQuery("SELECT ProjectNo from TblProject")
        .ClearForm
        .Show = True
    End With
    
    ActiveProject.DBGet Picker.SelectedItem

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
    
    ActiveLender.DBGet Picker.SelectedItem
    
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
    
    If Not FrmWFLender.ShowForm Then Err.Raise HANDLED_ERROR
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
    
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


