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

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen(ScrnPage As enScreenPage) As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ScreenPage = ScrnPage
    
    ShtMain.Unprotect PROTECT_KEY
    
    If Not BuildMainFrame Then Err.Raise HANDLED_ERROR
    If Not RefreshList Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
    
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
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    Set ButtonFrame = New ClsUIFrame
    Set BtnNewProjectWF = New ClsUIButton
    Set BtnNewLenderWF = New ClsUIButton
        
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    MainScreen.Frames.AddItem ButtonFrame, "Button Frame"
    
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
            .Text = "Active Projects"
            .Style = HEADER_STYLE
            .Visible = True
        End With

        With .Table
            .Left = GENERIC_TABLE_LEFT
            .Top = GENERIC_TABLE_TOP
            .HPad = GENERIC_TABLE_ROWOFFSET
            .VPad = GENERIC_TABLE_COLOFFSET
            .SubTableVOff = 50
            .SubTableHOff = 20
            .HeadingText = ACTIVE_TABLE_TITLES
            .HeadingStyle = GENERIC_TABLE_HEADER
            .HeadingHeight = GENERIC_TABLE_HEADING_HEIGHT
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
    
    With BtnNewProjectWF

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = ACTIVE_BTN_MAIN_1_LEFT
        .Top = ACTIVE_BTN_MAIN_1_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIButtons.ProcessBtnClicks(" & enBtnNewProjectWF & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Project Workflow"
    End With
    
    With BtnNewLenderWF

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = ACTIVE_BTN_MAIN_2_LEFT
        .Top = ACTIVE_BTN_MAIN_2_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain2"
        .OnAction = "'ModUIButtons.ProcessBtnClicks(" & enBtnNewLenderWF & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Lender Workflow"
    End With
    
    ButtonFrame.Buttons.Add BtnNewProjectWF
    ButtonFrame.Buttons.Add BtnNewLenderWF
    
    BuildMainFrame = True

Exit Function

ErrorExit:

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
Public Function RefreshList(Optional SortBy As String) As Boolean
    Dim NoCols As Integer
    Dim NoRows As Integer
    Dim Workflows As ClsWorkflows
    Dim StrSortBy As String
    Dim RstWorkflowList As Recordset
    Dim y As Integer
    Dim x As Integer
    Dim AryStyles() As String
    Dim AryOnAction() As String
    
    Const StrPROCEDURE As String = "RefreshList()"

    On Error GoTo ErrorHandler

    ModLibrary.PerfSettingsOn

    ShtMain.Unprotect PROTECT_KEY
    
    Set Workflows = New ClsWorkflows
    
    Set RstWorkflowList = GetActiveList(StrSortBy)
    
    With RstWorkflowList
        If .RecordCount = 0 Then GoTo GracefulExit
        
        .MoveLast
        .MoveFirst
    End With
    
    With MainFrame.Table
        .ColWidths = ACTIVE_TABLE_COL_WIDTHS
        .RstText = RstWorkflowList
        .NoRows = RstWorkflowList.RecordCount
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add GENERIC_TABLE_HEADER
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
                
                If x = 6 Then
                    If !RAG = "en1Red" Then AryStyles(x, y) = "RED_CELL"
                    If !RAG = "en2Amber" Then AryStyles(x, y) = "AMBER_CELL"
                    If !RAG = "en3Green" Then AryStyles(x, y) = "GREEN_CELL"
                Else
            AryStyles(x, y) = "GENERIC_TABLE"
                End If
                
                AryOnAction(x, y) = "'ModUIButtons.ProcessBtnClicks(""" & enBtnOpenProject & ":" & !ProjectNo & """)'"
                .MoveNext
        Next
    Next
    End With
    
    With MainFrame.Table
        .Styles = AryStyles
        .OnAction = AryOnAction
        .BuildCells
    End With
    
    ModLibrary.PerfSettingsOff

GracefulExit:
    
    RefreshList = True
 
    Workflows.Terminate
    Set RstWorkflowList = Nothing
    Set Workflows = Nothing
    
Exit Function

ErrorExit:

    Workflows.Terminate
    Set RstWorkflowList = Nothing
    Set Workflows = Nothing
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
' Method GetActiveList
' Gets data for workflow list
'---------------------------------------------------------------
Public Function GetActiveList(StrSortBy As String) As Recordset
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim SQL As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String

    SQL = "SELECT TblProject.ProjectNo, TblClient.Name, TblSPV.Name, TblCBSUser.UserName, TblWorkflow.CurrentStep, TblStepTemplate.StepName, TblWorkflow.Status, TblWorkflow.RAG " _
            & "FROM ((((TblProject LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblClient ON TblSPV.ClientNo = TblClient.ClientNo) LEFT JOIN TblCBSUser ON " _
            & "TblProject.CaseManager = TblCBSUser.CBSUserNo) LEFT JOIN TblWorkflow ON TblProject.ProjectWFNo = TblWorkflow.WorkflowNo) LEFT JOIN TblStepTemplate ON " _
            & "TblWorkflow.CurrentStep = TblStepTemplate.StepNo WHERE (((TblWorkflow.WorkflowType)='enProject'))"

                
    Set RstWorkflow = ModDatabase.SQLQuery(SQL)
    
    Set GetActiveList = RstWorkflow
    
End Function

