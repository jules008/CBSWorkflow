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
    
    MainFrame.ReOrder
       


    
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
       
'    Set RstRepData = ModDatabase.SQLQuery("SELECT COUNT(StudentID) AS [CountX] FROM TblRepData WHERE Active and QIP")
'    ArySource(1) = RstRepData![CountX]
'
'    Set RstRepData = ModDatabase.SQLQuery("SELECT COUNT(StudentID) AS [CountX] FROM TblRepData WHERE Active")
'    ArySource(2) = RstRepData![CountX]
    
'    With Graph1
'        .ChartType = enDoNut
'        .Name = "DNut1"
'        .DataLabels = False
'        MainFrame.Graphs.AddItem Graph1
'        .Height = GRAPH_1_HEIGHT
'        .Left = GRAPH_1_LEFT
'        .Top = GRAPH_1_TOP
'        .Ser1Colour = GRAPH_1_COL_1
'        .Ser2Colour = GRAPH_1_COL_2
'        .SourceData = ArySource
'        .Title = "Overall"
'        .GenGraph
'        .Visible = True
'
'    End With
'
'    Set RstRepData = ModDatabase.SQLQuery("SELECT COUNT(StudentID) AS [CountX] FROM TblRepData WHERE Active AND QIP AND Watch = 'Blue'")
'    ArySource(1) = RstRepData![CountX]
'
'    Set RstRepData = ModDatabase.SQLQuery("SELECT COUNT(StudentID) AS [CountX] FROM TblRepData WHERE Active AND Watch = 'Blue'")
'    ArySource(2) = RstRepData![CountX]
    
    BuildGraphs = True
        
Exit Function
    
ErrorExit:
    
    BuildGraphs = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

