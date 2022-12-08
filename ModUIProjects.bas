Attribute VB_Name = "ModUIProjects"
'===============================================================
' Module ModUIProjects
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 Jun 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIProjects"
Private ScreenPage As String
Private OldSplitRow As Integer
Private OldProjectNo As Integer

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen(ScrnPage As enScreenPage, ByRef SplitScreenOn As Boolean) As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ScreenPage = ScrnPage
    
    If Not BuildMainFrame(ScreenPage) Then Err.Raise HANDLED_ERROR
    
    If Not RefreshList(ScreenPage, SplitScreenOn) Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
    ModLibrary.PerfSettingsOff
                    
    BuildScreen = True
       
Exit Function

ErrorExit:
    
    ModLibrary.PerfSettingsOff

    BuildScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainFrame
' Builds screen components
' ---------------------------------------------------------------
Private Function BuildMainFrame(ByVal ScreenPage As enScreenPage, Optional SplitScreenInfo As String) As Boolean
    Dim HeaderText As String
    Dim TableHeadingText As String
    Dim TableColWidths As String
    Dim NewBtnNo As EnumBtnNo
    Dim NewBtnTxt As String
    Dim Contact As ClsContact
    
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    Set ButtonFrame = New ClsUIFrame
    Set BtnCommsToDo = New ClsUIButton
    Set BtnProjectNewWF = New ClsUIButton
    Set BtnNewLenderWF = New ClsUIButton
    Set Contact = New ClsContact
    
    With MainFrame.Table
        If Not .SubTable Is Nothing Then .SubTable.Terminate
        Set .SubTable = New ClsUITable
    End With
        
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    MainScreen.Frames.AddItem ButtonFrame, "Button Frame"
    
    'load page specific data
    Select Case ScreenPage
        Case enScrProjForAction
            TableHeadingText = PROJ_FOR_ACTION_HEADER_TEXT
        Case enScrProjActive
            TableHeadingText = PROJ_ACTIVE_HEADER_TEXT
        Case enScrProjComplete
            TableHeadingText = PROJ_CLOSED_HEADER_TEXT
    End Select
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
            
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = MAIN_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .ZOrder = 1

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
            .Text = TableHeadingText
            .Style = HEADER_STYLE
            .Visible = True
        End With

        With .Table
            .Left = GENERIC_TABLE_LEFT
            .Top = GENERIC_TABLE_TOP
            .HPad = GENERIC_TABLE_ROWOFFSET
            .VPad = GENERIC_TABLE_COLOFFSET
            .SubTableVOff = 50
            .SubTableHOff = 10
            .HeadingText = PROJECT_TABLE_TITLES
            .HeadingStyle = GENERIC_TABLE_HEADER
            .HeadingHeight = GENERIC_TABLE_HEADING_HEIGHT
            .ExpandIcon = GENERIC_TABLE_EXPAND_ICON
        End With
        
    End With
    
    With ButtonFrame
        .Top = BUTTON_FRAME_TOP
        .Left = BUTTON_FRAME_LEFT
        .Width = BUTTON_FRAME_WIDTH
        .Height = BUTTON_FRAME_HEIGHT
        .Style = BUTTON_FRAME_STYLE
        .EnableHeader = True
        .ZOrder = 1
        .Visible = False
    End With
    
    With BtnProjectNewWF

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = PROJECT_BTN_MAIN_1_LEFT
        .Top = PROJECT_BTN_MAIN_1_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnProjectNew & ":0" & """)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Project Workflow"
    End With
    
    With BtnNewLenderWF

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = PROJECT_BTN_MAIN_2_LEFT
        .Top = PROJECT_BTN_MAIN_2_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain2"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnLenderNewWF & ":0" & """)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Lender Workflow"
    End With
    
    With BtnCommsToDo

        .Height = GENERIC_BUTTON_HEIGHT
        .Width = TODO_BUTTON_WIDTH
        .Name = "BtnMain3"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnCommsToDo & ":0" & """)'"
        .UnSelectStyle = TODO_BUTTON
        .Selected = False
        .Text = "To Do Items"
        .Icon = ShtMain.Shapes.AddPicture(GetDocLocalPath(ThisWorkbook.Path) & PICTURES_PATH & TODO_ICON_FILE, msoTrue, msoFalse, 0, 0, 0, 0)
        With .Icon
            .Width = TODO_ICON_WIDTH
            .Height = TODO_ICON_WIDTH
            .Name = "ToDo Icon"
        End With
        .IconTop = TODO_ICON_TOP
        .IconLeft = TODO_ICON_LEFT
        .Badge = New ClsUICell
        With .Badge
            .Width = TODO_BADGE_WIDTH
            .Height = TODO_BADGE_HEIGHT
            .Name = "ToDo Badge"
            .Style = TODO_BADGE
            .Text = Contact.CommsNo
        End With
        .BadgeTop = TODO_BADGE_TOP
        .BadgeLeft = TODO_BADGE_LEFT
        .Left = TODO_BUTTON_LEFT
        .Top = TODO_BUTTON_TOP
    End With
    ButtonFrame.Buttons.Add BtnProjectNewWF
    ButtonFrame.Buttons.Add BtnNewLenderWF
    ButtonFrame.Buttons.Add BtnCommsToDo
    
    MainScreen.ReOrder
    
    Set Contact = Nothing
    BuildMainFrame = True

Exit Function

ErrorExit:

    Set Contact = Nothing
    BuildMainFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
' ===============================================================
' RefreshList
' Refreshes the list of active workflows
' ---------------------------------------------------------------
Public Function RefreshList(ByVal ScreenPage As enScreenPage, ByVal SplitScreenOn As Boolean, Optional SortBy As String) As Boolean
    Dim NoCols As Integer
    Dim NoRows As Integer
    Dim StrSortBy As String
    Dim RstWorkflowList As Recordset
    Dim y As Integer
    Dim x As Integer
    Dim AryStyles() As String
    Dim AryOnAction() As String
    
    Const StrPROCEDURE As String = "RefreshList()"

    On Error GoTo ErrorHandler

    ModLibrary.PerfSettingsOn

    Set RstWorkflowList = GetActiveList(ScreenPage, SortBy)
    
    With RstWorkflowList
        If .RecordCount = 0 Then GoTo GracefulExit
        
        .MoveLast
        .MoveFirst
    End With
    
    With MainFrame.Table
        .Name = "Table"
        .ColWidths = PROJECT_TABLE_COL_WIDTHS
        .RstText = RstWorkflowList
        .NoRows = RstWorkflowList.RecordCount
        .StylesColl.RemoveCollection
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add GENERIC_TABLE_HEADER
        .StylesColl.Add RED_CELL
        .StylesColl.Add AMBER_CELL
        .StylesColl.Add GREEN_CELL
        .RowHeight = GENERIC_TABLE_ROW_HEIGHT
    End With
    
    With MainFrame.Table.SubTable
            .HeadingText = PROJECT_SUB_TABLE_TITLES
            .HeadingStyle = SUB_TABLE_HEADER
            .HeadingHeight = GENERIC_TABLE_HEADING_HEIGHT
            .Name = "SubTable"
        .ColWidths = PROJECT_SUB_TABLE_COL_WIDTHS
        .StylesColl.RemoveCollection
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add SUB_TABLE_HEADER
        .StylesColl.Add RED_CELL
        .StylesColl.Add AMBER_CELL
        .StylesColl.Add GREEN_CELL
        .RowHeight = GENERIC_TABLE_ROW_HEIGHT
    End With
    
    NoRows = RstWorkflowList.RecordCount
    NoCols = MainFrame.Table.NoCols
    
    ReDim AryStyles(0 To NoCols - 1, 0 To NoRows - 1)
    ReDim AryOnAction(0 To NoCols - 1, 0 To NoRows - 1)
    
    MainFrame.Table.Cells.DeleteCollection
    Debug.Assert MainFrame.Table.Cells.Count = 0
    
    With RstWorkflowList
        For x = 0 To NoCols - 1
            .MoveFirst
            For y = 0 To NoRows - 1
                
                If y = 0 Then
                    'headers
                    AryOnAction(x, y) = "'ModUIProjects.SortBy (""" & ScreenPage & ":" & .Fields(x).Name & """)'"
                Else
                    If x = 0 Then
                        AryOnAction(x, y) = "'ModUIProjects.SplitScreen(""" & y + 1 & ":" & !ProjectNo & """)'"
                Else
                        AryOnAction(x, y) = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnProjectOpen & ":" & !ProjectNo & """)'"
                    End If
                End If
                
                If x = 9 Then
                    If !RAG = "en1Red" Then AryStyles(x, y) = "RED_CELL"
                    If !RAG = "en2Amber" Then AryStyles(x, y) = "AMBER_CELL"
                    If !RAG = "en3Green" Then AryStyles(x, y) = "GREEN_CELL"
                Else
                    AryStyles(x, y) = "GENERIC_TABLE"
                End If
                .MoveNext
        Next
    Next
    End With
    
    With MainFrame.Table
        .Styles = AryStyles
        .OnAction = AryOnAction
        .BuildTable
    End With
    
    If SplitScreenOn Then
        ModUIProjects.SplitScreen
    End If
    
    ModLibrary.PerfSettingsOff

GracefulExit:
    
    RefreshList = True
 
    Set RstWorkflowList = Nothing
    
    Exit Function

ErrorExit:

    Set RstWorkflowList = Nothing
    
    ModLibrary.PerfSettingsOff
    
    RefreshList = False

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
' SplitScreen
' Splits the screen and inserts the lender workflows
' ---------------------------------------------------------------
Public Sub SplitScreen(Optional NewRowInfo As String)
    Dim ErrNo As Integer
    Dim RstWorkflows As Recordset
    Dim SQL As String
    Dim AryRowInfo() As String
    Dim AryStyles() As String
    Dim AryOnAction() As String
    Dim NoCols As Integer
    Dim NoRows As Integer
    Dim y As Integer
    Dim x As Integer
    Dim SplitRow As Integer
    Dim ProjectNo As Integer
    Dim SQLSelect As String
    Dim SQLFrom As String
    Dim SQLWhere As String
    
    Const StrPROCEDURE As String = "SplitScreen()"

    On Error GoTo ErrorHandler
    
Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If NewRowInfo = "" Then
        'keeping split the same
        SplitRow = OldSplitRow
        ProjectNo = OldProjectNo
    Else
        'update with new info
        AryRowInfo = Split(NewRowInfo, ":")
    
        SplitRow = CInt(AryRowInfo(0))
    ProjectNo = CInt(AryRowInfo(1))
    End If
    
    If SplitRow = OldSplitRow And NewRowInfo <> "" Then
        MainFrame.Table.BuildTable 0
        
        OldSplitRow = 0
        MainScreen.ReOrder
    Else
        OldSplitRow = SplitRow
        OldProjectNo = ProjectNo
    
        SQLSelect = "SELECT TblWorkflow.WorkflowNo, " _
                        & "TblWorkflowType.DisplayName, " _
                        & "TblLender.Name, " _
                        & "TblWorkflow.CurrentStep, " _
                        & "TblStepTemplate.StepName, " _
                        & "TblWorkflow.Progress & '%' AS [Progress], " _
                        & "TblWorkflow.Status, " _
                        & "TblWorkflow.RAG "
       SQLFrom = "FROM (TblStepTemplate " _
                        & "RIGHT JOIN (TblWorkflow " _
                        & "LEFT JOIN TblLender ON TblWorkflow.LenderNo = TblLender.LenderNo) " _
                        & "ON TblStepTemplate.StepNo = TblWorkflow.CurrentStep) " _
                        & "LEFT JOIN TblWorkflowType ON TblWorkflow.Name = TblWorkflowType.WFName "
        SQLWhere = "WHERE (((TblWorkflow.ProjectNo)= " & ProjectNo & ") " _
                        & "AND ((TblWorkflow.WorkflowType)='enLender'))"

        SQL = SQLSelect & SQLFrom & SQLWhere

    Set RstWorkflows = ModDatabase.SQLQuery(SQL)
    
        If RstWorkflows.RecordCount > 0 Then
    With RstWorkflows
        .MoveLast
        .MoveFirst
        NoRows = .RecordCount
        NoCols = .Fields.Count
    End With
    
    With MainFrame.Table.SubTable
        .RstText = RstWorkflows
        .NoRows = RstWorkflows.RecordCount
    End With
    
    ReDim AryStyles(0 To NoCols - 1, 0 To NoRows - 1)
    ReDim AryOnAction(0 To NoCols - 1, 0 To NoRows - 1)
    
    With RstWorkflows
    
        For x = 0 To NoCols - 1
            .MoveFirst
            For y = 0 To NoRows - 1
                
                If x = 6 Then
                    If !RAG = "en1Red" Then AryStyles(x, y) = "RED_CELL"
                    If !RAG = "en2Amber" Then AryStyles(x, y) = "AMBER_CELL"
                    If !RAG = "en3Green" Then AryStyles(x, y) = "GREEN_CELL"
                Else
                    AryStyles(x, y) = "GENERIC_TABLE"
                End If
                
                            AryOnAction(x, y) = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnLenderOpenWF & ":" & !WorkflowNo & """)'"
                .MoveNext
            Next
        Next
    End With
    
        With MainFrame.Table.SubTable
            .Styles = AryStyles
            .OnAction = AryOnAction
        End With
    
            MainFrame.Table.BuildTable SplitRow, 100
        MainScreen.ReOrder
        Else
            MainFrame.Table.BuildTable 0
            MsgBox "There are no Lender Workflows for this Project", vbOKOnly + vbInformation, APP_NAME
        End If
    End If
    
GracefulExit:
    Set RstWorkflows = Nothing
    Exit Sub

ErrorExit:

    Set RstWorkflows = Nothing
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

' ===============================================================
' Method GetActiveList
' Gets data for workflow list
'---------------------------------------------------------------
Public Function GetActiveList(ScreenPage As enScreenPage, StrSortBy As String) As Recordset
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim SQL As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String

    If StrSortBy <> "" Then StrSortBy = "ORDER BY " & StrSortBy
    
    Select Case ScreenPage
        Case enScrProjForAction
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%' AS [Progress], TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.RAG)='en1Red') AND ((TblWorkflow.WorkflowType)='enProject')) OR (((TblWorkflow.RAG)='en2Amber') AND ((TblWorkflow.WorkflowType)='enProject')) OR (((TblWorkflow.Status)='Action Req.') AND ((TblWorkflow.RAG)='en3Green') AND ((TblWorkflow.WorkflowType)='enProject')) " _
                    & StrSortBy
        Case enScrProjActive
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%' AS [Progress], TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.Status)<>'Complete') AND ((TblWorkflow.WorkflowType)='enProject')) " _
                    & StrSortBy
        Case enScrProjComplete
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%' AS [Progress], TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.Status)='Complete') AND ((TblWorkflow.WorkflowType)='enProject')) " _
                    & StrSortBy
    End Select
                
    Set RstWorkflow = ModDatabase.SQLQuery(SQL)
    
    Set GetActiveList = RstWorkflow
    
End Function

' ===============================================================
' OpenProjectWF
' Opens project workflow
' ---------------------------------------------------------------
Public Function OpenProjectWF(ByVal ScreenPage As enScreenPage, ByVal Index As String) As Boolean
    Dim CRMItem As Object
    
    Const StrPROCEDURE As String = "OpenProjectWF()"

    On Error GoTo ErrorHandler

    Set ActiveProject = New ClsProject
    
    With ActiveProject
        .DBGet Index
        .ProjectWorkflow.DisplayForm
    End With
    
    Set ActiveProject = Nothing
    
    OpenProjectWF = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    OpenProjectWF = False

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
' OpenLenderWF
' Opens project workflow
' ---------------------------------------------------------------
Public Function OpenLenderWF(ByVal ScreenPage As enScreenPage, ByVal Index As String) As Boolean
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "OpenLenderWF()"

    On Error GoTo ErrorHandler
    
    Set RstResults = ModDatabase.SQLQuery("SELECT ProjectNo FROM TblWorkflow WHERE WorkflowNo = " & Index)
    
    Set ActiveWorkFlow = New ClsWorkflow
    Set ActiveProject = New ClsProject
    
    ActiveProject.DBGet RstResults!ProjectNo
    
    ActiveWorkFlow.DBGet Index
    ActiveProject.Workflows.Add ActiveWorkFlow
    ActiveWorkFlow.DisplayForm
    
    OpenLenderWF = True

    Set ActiveWorkFlow = Nothing
    Set ActiveProject = Nothing
    Set RstResults = Nothing
Exit Function

ErrorExit:

    OpenLenderWF = False

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
' SortBy
' Sorts cols by selected field
' ---------------------------------------------------------------
Private Sub SortBy(SortByData As String)
    Dim ArySort() As String
    Dim SortBy As String
    Dim StrSort As String
    Dim ScreenPage As enScreenPage
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "SortBy()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    ArySort = Split(SortByData, ":")
    ScreenPage = ArySort(0)
    SortBy = ArySort(1)
    
    Select Case SortBy
        Case Is = "ProjectNo"
            StrSort = "TblProject.ProjectNo"
        Case Is = "ProjectName"
            StrSort = "TblProject.ProjectName"
        Case Is = "TblClient.Name"
            StrSort = "TblClient.Name"
        Case Is = "TblSPV.Name"
            StrSort = "TblSPV.Name"
        Case Is = "UserName"
            StrSort = "TblCBSUser.UserName"
        Case Is = "CurrentStep"
            StrSort = "TblWorkflow.CurrentStep"
        Case Is = "StepName"
            StrSort = "TblStepTemplate.StepName"
        Case Is = "Progress"
            StrSort = "Progress"
        Case Is = "Status"
            StrSort = "TblWorkflow.Status"
    End Select

    If Not RefreshList(ScreenPage, False, StrSort) Then Err.Raise HANDLED_ERROR

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

