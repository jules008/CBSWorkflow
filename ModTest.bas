Attribute VB_Name = "ModTest"

    
Public Sub NewClasses()
    Dim Project As ClsProject
    Set CodeTimer = New ClsCodeTimer
    
    CodeTimer.StartTimer
    
    Set Project = New ClsProject
    
    CodeTimer.MarkTime "Project Class Created"
    
    With Project
        .DBGet 2
'        .ProjectWorkflow.Steps("1.01").Email.Body = "Red"
'        Debug.Print .ProjectWorkflow.Steps("1.01").Email.Body
'        .ProjectWorkflow.Steps("1.01").DisplayForm
'        .ProjectWorkflow.Steps("1.01").DisplayHelpForm
    End With
    CodeTimer.MarkTime "Project data retrieved from Database"
    

    Stop
    Project.Terminate
    Set Project = Nothing
End Sub

Public Sub TestAccessForm()
    Dim CBSUser As ClsCBSUser
    
    Set CBSUser = New ClsCBSUser
    
    CBSUser.DBGet 7
    
    FrmAccessCntrl.ShowForm CBSUser
    
    Set CBSUser = Nothing
End Sub
