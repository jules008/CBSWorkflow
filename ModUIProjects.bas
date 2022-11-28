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
Private SplitRow As Integer

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen(ScrnPage As enScreenPage) As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ScreenPage = ScrnPage
    
    If Not BuildMainFrame(ScreenPage) Then Err.Raise HANDLED_ERROR
    If Not RefreshList(ScreenPage) Then Err.Raise HANDLED_ERROR
    
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
Private Function BuildMainFrame(ByVal ScreenPage As enScreenPage) As Boolean
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
Public Function RefreshList(ByVal ScreenPage As enScreenPage, Optional SortBy As String) As Boolean
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

    Set RstWorkflowList = GetActiveList(ScreenPage, StrSortBy)
    
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
    
    Debug.Assert MainFrame.Table.Cells.Count = 0
    
    With RstWorkflowList
        For x = 0 To NoCols - 1
            .MoveFirst
            For y = 0 To NoRows - 1
                
                If x = 9 Then
                    If !RAG = "en1Red" Then AryStyles(x, y) = "RED_CELL"
                    If !RAG = "en2Amber" Then AryStyles(x, y) = "AMBER_CELL"
                    If !RAG = "en3Green" Then AryStyles(x, y) = "GREEN_CELL"
                Else
            AryStyles(x, y) = "GENERIC_TABLE"
                End If
                
                If x = 0 Then
                    AryOnAction(x, y) = "'ModUIProjects.SplitScreen(""" & y + 1 & ":" & !ProjectNo & """)'"
                Else
                AryOnAction(x, y) = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnProjectOpen & ":" & !ProjectNo & """)'"
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
Private Sub SplitScreen(RowInfo As String)
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
    Dim RowNo As Integer
    Dim ProjectNo As Integer
    
    Const StrPROCEDURE As String = "SplitScreen()"

    On Error GoTo ErrorHandler
    
Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    AryRowInfo = Split(RowInfo, ":")
    
    RowNo = CInt(AryRowInfo(0))
    ProjectNo = CInt(AryRowInfo(1))
    
    If SplitRow = RowNo Then
        MainFrame.Table.BuildTable 0
        
        SplitRow = 0
        MainScreen.ReOrder
    Else
    
    SQL = "SELECT TblWorkflow.WorkflowNo, TblLender.Name, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%',TblWorkflow.Status, TblWorkflow.RAG " _
                & "FROM TblStepTemplate RIGHT JOIN (TblWorkflow LEFT JOIN TblLender ON TblWorkflow.LenderNo = TblLender.LenderNo) ON TblStepTemplate.StepNo = TblWorkflow.CurrentStep " _
            & "WHERE (((TblWorkflow.ProjectNo)= " & ProjectNo & ") AND ((TblWorkflow.WorkflowType)='enLender'))"

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
                
                If x = 5 Then
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
    
        SplitRow = RowNo
        MainFrame.Table.BuildTable RowNo, 100
        MainScreen.ReOrder
        Else
            MainFrame.Table.BuildTable 0
        End If
    End If
    
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

    Select Case ScreenPage
        Case enScrProjForAction
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%', TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.RAG)='en1Red') AND ((TblWorkflow.WorkflowType)='enProject')) OR (((TblWorkflow.RAG)='en2Amber') AND ((TblWorkflow.WorkflowType)='enProject')) OR (((TblWorkflow.Status)='Action Req.') AND ((TblWorkflow.RAG)='en3Green') AND ((TblWorkflow.WorkflowType)='enProject'))"
        Case enScrProjActive
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%', TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.Status)<>'Complete') AND ((TblWorkflow.WorkflowType)='enProject'))"
        Case enScrProjComplete
            SQL = "SELECT Null AS Expand, TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Progress & '%', TblWorkflow.Status, TblWorkflow.RAG " _
                    & "FROM TblClient RIGHT JOIN ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblStepTemplate ON TblWorkflow.CurrentStep = TblStepTemplate.StepNo) ON TblClient.ClientNo = TblProject.ClientNo " _
                    & "WHERE (((TblWorkflow.Status)='Complete') AND ((TblWorkflow.WorkflowType)='enProject'))"
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
    Dim RstWorkflow As Recordset
    Dim CRMItem As Object
    
    Const StrPROCEDURE As String = "OpenLenderWF()"

    On Error GoTo ErrorHandler
    
    Set ActiveWorkFlow = New ClsWorkflow
    Set ActiveProject = New ClsProject
    
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT ProjectNo FROM TblWorkflow WHERE WorkflowNo = " & Index)
    
    With ActiveProject
        .DBGet RstWorkflow.Fields(0)
        .Workflows(Index).DisplayForm
    End With
    
    Set ActiveWorkFlow = Nothing
    Set ActiveProject = Nothing
    Set RstWorkflow = Nothing
    
    OpenLenderWF = True

Exit Function

ErrorExit:

    Set ActiveWorkFlow = Nothing
    Set ActiveProject = Nothing
    Set RstWorkflow = Nothing
    
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


