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

Public Sub ClientNeedTest()
    Dim ClientNeeds As Byte
    
    ClientNeeds = 255
    
    With FrmClientForm
        .SetClientNeed ClientNeeds
        .Show
        ClientNeeds = Format(.GetClientNeed, "0")
    End With
    Debug.Print ClientNeeds
End Sub

Public Sub TestPicker()
    Dim Picker As ClsFrmPicker
    
    Set Picker = New ClsFrmPicker
    With Picker
        .Title = "Select workflow script"
        .Instructions = "Select the workflow script you would like to view."
        .Data = ModDatabase.SQLQuery("SELECT SecondTier from TblWorkflowTable")
        .ClearForm
        .Show = True
    End With
    
    
    Set Picker = Nothing
    
End Sub

Public Sub TestCellCls()
    Dim Cell As ClsUICell
    Dim Shp As Shape
    
    Set Shp = ShtMain.Shapes.AddPicture(GetDocLocalPath(ThisWorkbook.Path) & PICTURES_PATH & TODO_ICON_FILE, msoTrue, msoFalse, 0, 0, 0, 0)
    Shp.Name = "temp"
    Set Cell = New ClsUICell
    
    With Cell
        .Left = 100
        .Top = 100
        .Height = 500
        .Width = 1000
    End With
    
    Cell.Badges.Add Shp
    
    With Cell.Badges
        .SetLeft(Shp.Name) = 100
        .SetTop(Shp.Name) = 200
        .SetHeight(Shp.Name) = 100
        .SetWidth(Shp.Name) = 100
    End With
    
    Cell.ReOrder
    Stop
    Set Shp = Nothing
    Set Shp = Cell.Badges("temp")
    Shp.Delete
    Set Shp = Nothing
    Cell.Terminate
    Set Cell = Nothing
End Sub
