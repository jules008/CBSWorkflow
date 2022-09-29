Attribute VB_Name = "ModUIActive"
'===============================================================
' Module ModUIActive
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

Private Const StrMODULE As String = "ModUIActive"

' ===============================================================
' BuildMainFrame
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
            
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = MAIN_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
            .Text = "Active Workflows"
            .Style = HEADER_STYLE
            .Visible = True
        End With

        With .Table
            .Left = GENERIC_TABLE_LEFT
            .Top = GENERIC_TABLE_TOP
            .NoCols = ACTIVE_TABLE_NOCOLS
            .HPad = GENERIC_TABLE_ROWOFFSET
            .VPad = GENERIC_TABLE_COLOFFSET
            .SubTableVOff = 50
            .SubTableHOff = 20
            .HeadingText = ACTIVE_TABLE_TITLES
            .HeadingStyle = GENERIC_TABLE_HEADER
            .HeadingHeight = GENERIC_TABLE_HEADING_HEIGHT
            .ColWidths = ACTIVE_TABLE_COL_WIDTHS
            .RowHeights = GENERIC_TABLE_HEIGHT
        End With
    End With
    
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
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ShtMain.Unprotect PROTECT_KEY
    
    If Not ModUIActive.BuildMainFrame Then Err.Raise HANDLED_ERROR
    If Not ModUIScreenCom.BuildScreenBtn1 Then Err.Raise HANDLED_ERROR
    If Not ModUIActive.RefreshList Then Err.Raise HANDLED_ERROR
    
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
' RefreshList
' Refreshes the list of active workflows
' ---------------------------------------------------------------
Public Function RefreshList(Optional SortBy As String) As Boolean
'    Dim StepNo As String
'    Dim CurrentStep As String
'    Dim ActionOn As String
'    Dim StepStatus As enStatus
    Dim NoCols As Integer
    Dim NoRows As Integer
    Dim Workflows As ClsWorkflows
'    Dim StrOnAction As String
    Dim StrSortBy As String
    Dim RstWorkflowList As Recordset
'    Dim ScreenSel As String
    Dim y As Integer
    Dim x As Integer
'    Dim RowTitles() As String
'    Dim WorkflowNo As String
'    Dim MemberName As String
    Dim AryStyles() As String

    Const StrPROCEDURE As String = "RefreshList()"

    On Error GoTo ErrorHandler

    ModLibrary.PerfSettingsOn

    ShtMain.Unprotect PROTECT_KEY
    
'    ScreenSel = "UIActive"
    
    Set Workflows = New ClsWorkflows
    Workflows.UpdateRAGs
    
    Set RstWorkflowList = GetActiveList(StrSortBy)
    
    With RstWorkflowList
        .MoveLast
        .MoveFirst
        If .RecordCount = 0 Then GoTo GracefulExit
            
    End With
    
    NoCols = ACTIVE_TABLE_NOCOLS
    NoRows = RstWorkflowList.RecordCount + 1
    
    ReDim AryStyles(1 To NoCols, 0 To NoRows)
    
    For x = 1 To NoCols
        For y = 0 To NoRows
            AryStyles(x, y) = "GENERIC_TABLE"
        Next
    Next
    
    With MainFrame.Table
        .RstText = RstWorkflowList
        .NoRows = RstWorkflowList.RecordCount
        .Styles = AryStyles
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add GENERIC_TABLE_HEADER
        .RowHeights = FOR_ACTION_CELL_ROW_HEIGHTS
        .ColWidths = FOR_ACTION_CELL_COL_WIDTHS
        .BuildCells
    End With
    
'    With MainFrame
'        For Each Lineitem In .Lineitems
'            .Lineitems.RemoveItem Lineitem.Name
'            Lineitem.ShpLineitem.Delete
'            Set Lineitem = Nothing
'        Next
'
'        ReDim RowTitles(0 To ACTIVE_LINEITEM_NOCOLS - 1)
'        RowTitles = Split(ACTIVE_LINEITEM_TITLES, ":")
'
'        For i = 0 To ACTIVE_LINEITEM_NOCOLS - 1
'            StrOnAction = "'ModUIScreenCom.SortBy(""" & RowTitles(i) & """), """ & ScreenSel & """'"
'            .Lineitems.Text 0, i, RowTitles(i), GENERIC_TABLE_HEADER, StrOnAction
'        Next
'
'        .Lineitems.Style = GENERIC_TABLE
'
'    End With
    
'    x = 1
    
    If SortBy = "" Then
        If StrSortBy = "" Then
            StrSortBy = "TblMember.DisplayName"
        End If
    Else
       StrSortBy = SortBy
    End If
    
            
'    With RstWorkflowList
'        .MoveLast
'        .MoveFirst
'        For x = 1 To .RecordCount
'
'            StrOnAction = "'ModUIActive.OpenWorkflow(" & !WorkflowNo & ")'"
'
'            If Not IsNull(!CurrentStep) Then StepNo = !CurrentStep Else StepNo = ""
'            If Not IsNull(!WorkflowNo) Then WorkflowNo = !WorkflowNo Else WorkflowNo = ""
'            If Not IsNull(!Member) Then MemberName = !Member Else MemberName = ""
'            If Not IsNull(!StepName) Then CurrentStep = !StepName Else CurrentStep = ""
'            ActionOn = ""
'            If Not IsNull(!Status) Then StepStatus = enStatusVal(!Status)
        
'            If Not IsNull(!RAG) Then
'                Select Case enRAGVal(!RAG)
'                    Case Is = en3Green
'                        CustomStyle = GREEN_LINEITEM
'                    Case en2Amber
'                        CustomStyle = AMBER_LINEITEM
'                    Case Is = en1Red
'                        CustomStyle = RED_LINEITEM
'                End Select
'            End If
            
'            With MainFrame.Lineitems
'                .Text x, 0, WorkflowNo, GENERIC_TABLE, StrOnAction
'                .Text x, 1, MemberName, GENERIC_TABLE, StrOnAction
'                .Text x, 2, StepNo, GENERIC_TABLE, StrOnAction
'                .Text x, 3, CurrentStep, GENERIC_TABLE, StrOnAction
'                .Text x, 4, enStatusDisp(StepStatus), CustomStyle, StrOnAction
'            End With
'
'            If x > ACTIVE_MAX_LINES Then Exit For
'
'            .MoveNext
'        Next
'    End With
    
'    MenuBar.Menu(1).BadgeText = Workflows.CountForAction
    
'    MainFrame.Height = (x + 1) * 21
'    If MainScreen.Height < MainFrame.Height + 500 Then
'        MainScreen.Height = MainFrame.Height + 500
'    End If
    
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
' OpenWorkflow
' Opens the selected Workflow
' ---------------------------------------------------------------
Private Sub OpenWorkflow(WorkflowNo As Integer)
    
    Const StrPROCEDURE As String = "OpenWorkflow()"
       
    On Error GoTo ErrorHandler

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
        
Restart:
'    If CurrentUser.UserLvl = enBasic Then Err.Raise ACCESS_DENIED
    
    Set ActiveWorkFlow = Nothing
    Set ActiveWorkFlow = New ClsWorkflow
    ActiveWorkFlow.DBGet CStr(WorkflowNo)
    
    If Not FrmWorkflow.ShowForm() Then Err.Raise HANDLED_ERROR
   
    If Not ModUIActive.RefreshList Then Err.Raise HANDLED_ERROR
    
GracefulExit:
Exit Sub

ErrorExit:
    ModLibrary.PerfSettingsOff
    Terminate
Exit Sub

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        If CustomErrorHandler(Err.Number) = SYSTEM_RESTART Then
            Resume Restart
        Else
            Resume GracefulExit
        End If
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
Public Function GetActiveList(StrSortBy As String) As Recordset
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim SQL As String
    Dim i As Integer

    SQL = ("SELECT " _
                & "TblWorkflow.Name, " _
                & "TblWorkflow.WorkflowNo, " _
                & "TblWorkflow.CurrentStep, " _
                & "TblWorkflow.Status, " _
                & "TblWorkflow.Member, " _
                & "TblStep.StepName, " _
                & "TblWorkflow.RAG, " _
                & "TblWorkflow.WorkflowNo " _
            & "FROM " _
                & "TblWorkflow " _
            & "INNER JOIN TblStep ON TblStep.StepNo = TblWorkflow.CurrentStep AND TblStep.WorkflowNo = TblWorkflow.WorkflowNo " _
            & "WHERE " _
                & "TblWorkflow.Status <> 'enComplete'")
                
    Set RstWorkflow = ModDatabase.SQLQuery(SQL)
    
    Set GetActiveList = RstWorkflow
    
End Function

