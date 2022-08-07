Attribute VB_Name = "ModUIComplete"
'===============================================================
' Module ModUIComplete
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

Private Const StrMODULE As String = "ModUIComplete"

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
            .Text = "Completed Workflows"
            .Style = HEADER_STYLE
'            .Icon = ShtMain.Shapes("TEMPLATE - Complete").Duplicate
'            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
'            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
'            .Icon.Name = .Parent.Name & " Icon"
'            .Icon.Visible = msoCTrue
        End With

        With .Lineitems
            .NoColumns = COMPLETE_LINEITEM_NOCOLS
            .Top = GENERIC_LINEITEM_TOP
            .Left = GENERIC_LINEITEM_LEFT
            .Height = GENERIC_LINEITEM_HEIGHT
            .Columns = COMPLETE_LINEITEM_COL_WIDTHS
            .RowOffset = GENERIC_LINEITEM_ROWOFFSET
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
' BuildScreenBtn1
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildScreenBtn1() As Boolean

    Const StrPROCEDURE As String = "BuildScreenBtn1()"

    On Error GoTo ErrorHandler

    Set BtnNewWorkflow = New ClsUIMenuItem

    With BtnNewWorkflow
        
        .Height = BTN_MAIN_1_HEIGHT
        .Left = BTN_MAIN_1_LEFT
        .Top = BTN_MAIN_1_TOP
        .Width = BTN_MAIN_1_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIComplete.ProcessBtnPress(" & enBtnNewWorkflow & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Workflow"
    End With

    MainFrame.Menu.AddItem BtnNewWorkflow
    
    BuildScreenBtn1 = True

Exit Function

ErrorExit:

    BuildScreenBtn1 = False

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
    
    If Not ModUIComplete.BuildMainFrame Then Err.Raise HANDLED_ERROR
'    If Not ModUIComplete.BuildScreenBtn1 Then Err.Raise HANDLED_ERROR
    If Not ModUIComplete.RefreshList Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
'    ShtMain.Shapes("TEMPLATE - Reset").ZOrder msoSendToBack
'    ShtMain.Shapes("TEMPLATE - Logo II").ZOrder msoSendToBack
    
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
' DestroyMainScreen
' Destroys the main screen objects
' ---------------------------------------------------------------
Public Function DestroyMainScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "DestroyMainScreen()"

    On Error GoTo ErrorHandler
    
    Set Frame = New ClsUIFrame
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
        End If
    Next
        
    If Not MainFrame Is Nothing Then MainFrame.Visible = False
    
    Set MainFrame = Nothing
    
    Set Frame = Nothing
    
    DestroyMainScreen = True
       
Exit Function

ErrorExit:

    Set Frame = Nothing
    
    DestroyMainScreen = False
    
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
' Refreshes the list of Complete workflows
' ---------------------------------------------------------------
Public Function RefreshList() As Boolean
    Dim StepNo As String
    Dim CurrentStep As String
    Dim ActionOn As String
    Dim StepStatus As enStatus
    Dim Workflows As ClsWorkflows
    Dim Lineitem As ClsUILineitem
    Dim StrOnAction As String
    Dim CustomStyle As TypeStyle
    Dim StrSortBy As String
    Dim RstCompWorkflow As Recordset
    Dim ScreenSel As String
    Dim i As Integer
    Dim x As Integer
    Dim RowTitles() As String
    Dim WorkflowNo As String
    Dim MemberName As String

    Const StrPROCEDURE As String = "RefreshList()"

    On Error GoTo ErrorHandler

    ModLibrary.PerfSettingsOn

    ShtMain.Unprotect PROTECT_KEY
    
    With MainFrame
        For Each Lineitem In .Lineitems
            .Lineitems.RemoveItem Lineitem.Name
            Set Lineitem = Nothing
        Next

        ReDim RowTitles(0 To COMPLETE_LINEITEM_NOCOLS - 1)
        RowTitles = Split(COMPLETE_LINEITEM_TITLES, ":")

        For i = 0 To COMPLETE_LINEITEM_NOCOLS - 1
            .Lineitems.Text 0, i, RowTitles(i), GENERIC_LINEITEM_HEADER, False
        Next
        
    End With
    
    x = 1
    
    Set RstCompWorkflow = ModDatabase.SQLQuery("SELECT " _
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
                                                & "TblWorkflow.Status = 'enComplete' And " _
                                                & "TblWorkflow.Deleted Is Null ")
    
    With RstCompWorkflow
        Debug.Print .RecordCount
        Do While Not .EOF
        
            StrOnAction = ""
            
            If Not IsNull(!CurrentStep) Then StepNo = !CurrentStep Else StepNo = ""
            If Not IsNull(!WorkflowNo) Then WorkflowNo = !WorkflowNo Else WorkflowNo = ""
            If Not IsNull(!Member) Then MemberName = !Member Else MemberName = ""
            If Not IsNull(!StepName) Then CurrentStep = !StepName Else CurrentStep = ""
            ActionOn = ""
            If Not IsNull(!Status) Then StepStatus = enStatusVal(!Status)
            
            With MainFrame.Lineitems
                .Text x, 0, WorkflowNo, GENERIC_LINEITEM, StrOnAction
                .Text x, 1, MemberName, GENERIC_LINEITEM, StrOnAction
                .Text x, 2, StepNo, GENERIC_LINEITEM, StrOnAction
                .Text x, 3, CurrentStep, GENERIC_LINEITEM, StrOnAction
                .Text x, 4, enStatusDisp(StepStatus), CustomStyle, StrOnAction
            End With
            
            If x > COMPLETE_MAX_LINES Then Exit Do
            
            x = x + 1
            .MoveNext
        Loop
    End With
    
'    MenuBar.Menu(1).BadgeText = Workflows.CountForAction
    
    ModLibrary.PerfSettingsOff

    RefreshList = True
    Set RstCompWorkflow = Nothing
    
Exit Function

ErrorExit:

    Set RstCompWorkflow = Nothing
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
   
    If Not ModUIComplete.RefreshList Then Err.Raise HANDLED_ERROR
    
    ActiveWorkFlow.DBSave
    
GracefulExit:

Exit Sub

ErrorExit:
    ModLibrary.PerfSettingsOff
    Terminate
Exit Sub

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        If CustomErrorHandler(Err.Number) = SYSTEM_RESTART Then
            BuildScreen
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

