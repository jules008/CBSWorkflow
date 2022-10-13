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
' ---------------------------------------------------------------
Public Sub BtnProjectNewWFClick()
    Dim Picker As ClsFrmPicker
    
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
    
    Set ActiveWorkFlow = Nothing
    Set ActiveSPV = Nothing
    Set ActiveProject = Nothing
    Set Picker = Nothing
    Set ActiveClient = Nothing
    Set ActiveUser = Nothing
    
'    If Not ModUIProjects.BuildScreen(ScreenPage) Then Err.Raise HANDLED_ERROR
End Sub

' ===============================================================
' BtnProjectOpenWFClick
' ---------------------------------------------------------------
Public Sub BtnProjectOpenWFClick(ByVal ScreenPage As enScreenPage, ByVal Index As String)
    If Not ModUIProjects.OpenItem(ScreenPage, Index) Then Err.Raise HANDLED_ERROR
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


