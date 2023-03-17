Attribute VB_Name = "ModUIDashboard"
'===============================================================
' Module ModUIDashboard
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 Dec 22
'===============================================================												  						  

Option Explicit

Private Const StrMODULE As String = "ModUIDashboard"



' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
 
    
    Application.ScreenUpdating = False
    
    If Not BuildMainFrame Then Err.Raise HANDLED_ERROR
    If Not BuildGraphs Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
    Application.ScreenUpdating = True
    
    If Not DEV_MODE Then ShtMain.Protect PROTECT_KEY
    
    ModLibrary.PerfSettingsOff
                    
    BuildScreen = True
       
Exit Function

ErrorExit:
    
    Application.ScreenUpdating = True
    
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
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = DASHBOARD_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
'        .Lineitems.Style = GENERIC_LINEITEM
		.ZOrder = 1
        
        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
            .Text = "Dashboard"
            .Style = HEADER_STYLE
            .Visible = True
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
' BuildGraphs
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildGraphs() As Boolean
    Dim RstRepData As Recordset
    Dim ArySource() As Variant
    Dim TotalNo As Integer
    Dim SQL As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BuildGraphs()"

    On Error GoTo ErrorHandler
    
    Set Graph1 = New ClsUIGraph
    Set Graph2 = New ClsUIGraph
    Set Graph3 = New ClsUIGraph
    Set Graph4 = New ClsUIGraph
    Set Graph5 = New ClsUIGraph
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
        
    Set RstRepData = GetGraphData(1)
    
    With RstRepData
        .MoveLast
        .MoveFirst
        ArySource = RstRepData.GetRows(.RecordCount)
    End With
    
    With Graph1
        .ChartType = enBarStacked
        .Name = "Graph1"
        .DataLabels = True
        MainFrame.Graphs.AddItem Graph1
        .Height = GRAPH_1_HEIGHT
        .Left = GRAPH_1_LEFT
        .Top = GRAPH_1_TOP
        .Ser1Colour = GRAPH_1_COL_1
        .Ser2Colour = GRAPH_1_COL_2
        .Ser1Name = "Closed"
        .Ser2Name = "Open"
        .BackColour = GRAPH_1_BACK_COL
        .SourceData = ArySource
        .Title = "Cases Open/Closed"
        .GenGraph
        .Visible = True

    End With
    
    Set RstRepData = GetGraphData(2)
    
    With RstRepData
        .MoveLast
        .MoveFirst
        ArySource = RstRepData.GetRows(.RecordCount)
    End With
    
    With Graph2
        .ChartType = enPie
        .Name = "Graph2"
        .DataLabels = True
        MainFrame.Graphs.AddItem Graph2
        .Height = GRAPH_2_HEIGHT
        .Left = GRAPH_2_LEFT
        .Top = GRAPH_2_TOP
        .Ser1Colour = GRAPH_2_COL_1
        .Ser2Colour = GRAPH_2_COL_2
        .Ser3Colour = GRAPH_2_COL_3
        .Ser4Colour = GRAPH_2_COL_4
        .Ser5Colour = GRAPH_2_COL_5
        .Ser6Colour = GRAPH_2_COL_6
        .BackColour = GRAPH_2_BACK_COL
        .SourceData = ArySource
        .Title = "Case Manager Cases"
        .GenGraph
        .Visible = True

    End With
    
    Set RstRepData = GetGraphData(3)
    
    With RstRepData
        .MoveLast
        .MoveFirst
        ArySource = RstRepData.GetRows(.RecordCount)
    End With
    
    With Graph3
        .ChartType = enPie
        .Name = "Graph3"
        .DataLabels = True
        MainFrame.Graphs.AddItem Graph3
        .Height = GRAPH_3_HEIGHT
        .Left = GRAPH_3_LEFT
        .Top = GRAPH_3_TOP
        .Ser1Colour = GRAPH_3_COL_1
        .Ser2Colour = GRAPH_3_COL_2
        .Ser3Colour = GRAPH_3_COL_3
        .Ser4Colour = GRAPH_3_COL_4
        .Ser5Colour = GRAPH_3_COL_5
        .Ser6Colour = GRAPH_3_COL_6
        .BackColour = GRAPH_3_BACK_COL
        .SourceData = ArySource
        .Title = "Client Introducer Cases"
        .GenGraph
        .Visible = True

    End With
    
    Set RstRepData = GetGraphData(4)
    
    With RstRepData
        .MoveLast
        .MoveFirst
        ArySource = RstRepData.GetRows(.RecordCount)
    End With
    
    With Graph4
        .ChartType = enline
        .Name = "Graph4"
        .DataLabels = False
        MainFrame.Graphs.AddItem Graph4
        .Height = GRAPH_4_HEIGHT
        .Left = GRAPH_4_LEFT
        .Top = GRAPH_4_TOP
        .Ser1Colour = GRAPH_4_COL_1
        .Ser2Colour = GRAPH_4_COL_2
        .Ser3Colour = GRAPH_4_COL_3
        .Ser1Name = "Dev."
        .Ser2Name = "Bridge"
        .Ser3Name = "Comm."
        .BackColour = GRAPH_4_BACK_COL
        .SourceData = ArySource
        .Title = "Average Loan Completion Time"
        .GenGraph
        .Visible = True

    End With
    
    Set RstRepData = GetGraphData(5)
    
    With RstRepData
        .MoveLast
        .MoveFirst
        ArySource = RstRepData.GetRows(.RecordCount)
    End With
    
    With Graph5
        .ChartType = enDoNut
        .Name = "Graph5"
        .DataLabels = True
        MainFrame.Graphs.AddItem Graph5
        .Height = GRAPH_5_HEIGHT
        .Left = GRAPH_5_LEFT
        .Top = GRAPH_5_TOP
        .Ser1Colour = GRAPH_5_COL_1
        .Ser2Colour = GRAPH_5_COL_2
        .Ser3Colour = GRAPH_5_COL_3
        .BackColour = GRAPH_5_BACK_COL
        .SourceData = ArySource
        .Title = "Number of Cases"
        .GenGraph
        .Visible = True

    End With
    
    
    Set RstRepData = Nothing
    BuildGraphs = True
        
Exit Function
    
ErrorExit:
    
    Set RstRepData = Nothing
    BuildGraphs = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Method GetGraphData
' Gets data for workflow list
'---------------------------------------------------------------
Public Function GetGraphData(GraphNo As String) As Recordset
    Dim RstData1 As Recordset
    Dim RstData2 As Recordset
    Dim SQL1 As String
    Dim SQL2 As String
    Dim Query1 As QueryDef
    Dim Query2 As QueryDef
    Dim ResultData As String
    
    Select Case GraphNo
        Case 1
            SQL1 = "Select " _
                & "    Format (TblTrendData.DataDate, 'dd mmm'), " _
                & "    TblTrendData.Closed, " _
                & "    TblTrendData.[Open] " _
                & "From " _
                & "    TblTrendData " _
                & "Where " _
                & "    TblTrendData.DataDate > DateAdd('d', Now(), -7) " _
                & "Order By " _
                & "    TblTrendData.DataDate ASc "
            
        Case 2
            SQL1 = "Select " _
                & "    * " _
                & "From " _
                & "    [CM Cases] " _
                & "Where  " _
                & "    CaseManager is not null and " _
                & "    NoCases > 0 "
                
        Case 3
            SQL1 = "Select " _
                & "    * " _
                & "From " _
                & "    [CI Cases] " _
                & "Where  " _
                & "    ClientIntroducer is not null and " _
                & "    NoCases > 0 "
        
        Case 4
            SQL1 = "Select " _
                & "    TblTrendData.DataDate, " _
                & "    TblTrendData.AveDev, " _
                & "    TblTrendData.AveBridge, " _
                & "    TblTrendData.AveComm " _
                & "From " _
                & "    TblTrendData "
                
        Case 5
            SQL1 = "Select " _
                & "    TblWorkflow.LoanType, " _
                & "    Count(TblWorkflow.WorkflowNo) As [Count_WorkflowNo] " _
                & "From " _
                & "    TblWorkflow " _
                & "Where " _
                & "    TblWorkflow.WorkflowType = 'enLender' And " _
                & "    TblWorkflow.LoanType Is Not Null And " _
                & "    TblWorkflow.Status <> 'Complete' " _
                & "Group By " _
                & "    TblWorkflow.LoanType "
            
    End Select
    
    Set RstData1 = ModDatabase.SQLQuery(SQL1)
    Set GetGraphData = RstData1
    
    Set RstData1 = Nothing
    Set RstData2 = Nothing
    Set Query1 = Nothing
    Set Query2 = Nothing
    
    
End Function

' ===============================================================
' UpdateTrendData
' Updates the trend data for the reports
' ---------------------------------------------------------------
Public Function UpdateTrendData() As Boolean
    Dim RstData1 As Recordset
    Dim RstData2 As Recordset
    Dim RstTrendTable As Recordset
    Dim RstAveTimes As Recordset
    Dim RstMaxDate As Recordset
    Dim MaxDate As Date
    Dim AddDate As Date
    Dim CullDate As Date
    Dim EditFlag As Boolean
    
    Const StrPROCEDURE As String = "UpdateTrendData()"

    On Error GoTo ErrorHandler

    Set RstTrendTable = ModDatabase.SQLQuery("TblTrendData")
    Set RstAveTimes = ModDatabase.SQLQuery("ProjTimeAve")
    Set RstMaxDate = ModDatabase.SQLQuery("Select Max(DataDate) as MaxDate FROM TblTrendData")
    Set RstData1 = ModDatabase.SQLQuery("SELECT  " _
                                     & "  Active.Active, " _
                                     & "  Closed.Closed  " _
                                     & "from  " _
                                     & "  Active, Closed ")
            
    With RstMaxDate
        If Not IsNull(!MaxDate) Then
            MaxDate = !MaxDate
        Else
            MaxDate = DateAdd("d", -1, Now())
        End If
    End With
    
    With RstTrendTable
        If DatePart("d", MaxDate) = DatePart("d", Now()) Then
            EditFlag = True
        Else
            EditFlag = False
            AddDate = DateAdd("d", 1, MaxDate)
        End If
        
        Do While DatePart("d", AddDate) <= DatePart("d", Now()) Or EditFlag
        
            If EditFlag Then
            .FindFirst "datadate = #" & MaxDate & "#"
            .Edit
                EditFlag = False
        Else
                .AddNew
                !DataDate = Format(AddDate, "dd mmm yy")
            End If
        
                !Open = RstData1!Active
                !Closed = RstData1!Closed
                
                With RstAveTimes
                    If .RecordCount > 0 Then
                    .MoveFirst
                    .FindFirst "Avg_LoanType = 'Commercial Mortgage'"
                    If Not .NoMatch Then RstTrendTable!AveComm = !NoDays
                    
                    .MoveFirst
                .FindFirst ("Avg_LoanType = 'Development Finance'")
                    If Not .NoMatch Then RstTrendTable!AveDev = !NoDays
                    
                    .MoveFirst
                    .FindFirst ("Avg_LoanType = 'Bridge/ Exit Loan'")
                    If Not .NoMatch Then RstTrendTable!AveBridge = !NoDays
                    End If
                End With
                .Update
                AddDate = DateAdd("d", 1, AddDate)
            Loop
    End With
    
    CullDate = DateAdd("d", 0 - GRAPH_TREND_DAYS, Now())
    
    DB.Execute "DELETE FROM TblTrendData WHERE DataDate < #" & CullDate & "#"
    
    Set RstData1 = Nothing
    Set RstData2 = Nothing
    Set RstTrendTable = Nothing
    Set RstAveTimes = Nothing
    Set RstMaxDate = Nothing

    UpdateTrendData = True


Exit Function

ErrorExit:

    Set RstData1 = Nothing
    Set RstData2 = Nothing
    Set RstTrendTable = Nothing
    Set RstAveTimes = Nothing
    Set RstMaxDate = Nothing
    
    UpdateTrendData = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

