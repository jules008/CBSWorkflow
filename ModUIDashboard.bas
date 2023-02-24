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
        .Height = MAIN_FRAME_HEIGHT
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
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
        
    ReDim ArySource(1 To 2)
       
    Set RstRepData = GetGraphData(1)
    
    With RstRepData
        ArySource(1) = !Complete
        ArySource(2) = !NotComplete
    End With
    
    With Graph1
        .ChartType = enDoNut
        .Name = "Graph1"
        .DataLabels = True
        MainFrame.Graphs.AddItem Graph1
        .Height = GRAPH_1_HEIGHT
        .Left = GRAPH_1_LEFT
        .Top = GRAPH_1_TOP
        .Ser1Colour = GRAPH_1_COL_1
        .Ser2Colour = GRAPH_1_COL_2
        .BackColour = GRAPH_1_BACK_COL
        .SourceData = ArySource
        .Title = "Cases Open/Closed"
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
            Set RstData1 = ModDatabase.SQLQuery("SELECT  " _
                                             & "  Active.TtlActive, " _
                                             & "  Closed.TtlClosed  " _
                                             & "from  " _
                                             & "  Active, Closed ")
            
        Case 2
                
        Case 3
        
        Case 4
            
    End Select
    
    Set RstData1 = Nothing
    Set RstData2 = Nothing
    Set Query1 = Nothing
    Set Query2 = Nothing
    
Set GetGraphData = RstData1
    
End Function


