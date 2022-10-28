Attribute VB_Name = "ModUICRM"
'===============================================================
' Module ModUICRM
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

Private Const StrMODULE As String = "ModUICRM"
Private ScreenPage As String

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen(ByVal ScrnPage As enScreenPage, Optional Filter As String) As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ScreenPage = ScrnPage
    
    If Not BuildMainFrame(ScreenPage) Then Err.Raise HANDLED_ERROR
    If Not RefreshList(ScreenPage, , Filter) Then Err.Raise HANDLED_ERROR
    
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
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame(ByVal ScreenPage As enScreenPage) As Boolean
    Dim HeaderText As String
    Dim TableHeadingText As String
    Dim TableColWidths As String
    Dim NewBtnNo As EnumBtnNo
    Dim NewBtnTxt As String
    Dim LeadBtnVisible As Boolean
    Dim CalImpBtnVisible As Boolean
    
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    Set ButtonFrame = New ClsUIFrame
    Set BtnCRMNewItem = New ClsUIButton
    Set BtnCRMContCalImp = New ClsUIButton
    Set BtnCRMContShwLead = New ClsUIButton
        
    MainScreen.Frames.AddItem MainFrame, "Main Frame"
    MainScreen.Frames.AddItem ButtonFrame, "Button Frame"
    
    'load page specific data
    Select Case ScreenPage
        Case enScrCRMClient
        
            HeaderText = "CRM - Clients"
            TableHeadingText = CRM_CLIENT_TABLE_TITLES
            TableColWidths = CRM_CLIENT_TABLE_COL_WIDTHS
            NewBtnTxt = "New Client"
            
        Case enScrCRMSPV
        
            HeaderText = "CRM - SPVs"
            TableHeadingText = CRM_SPV_TABLE_TITLES
            TableColWidths = CRM_SPV_TABLE_COL_WIDTHS
            NewBtnTxt = "New SPV"
            
        Case enScrCRMContact
        
            HeaderText = "CRM - Contacts"
            TableHeadingText = CRM_CONTACT_TABLE_TITLES
            TableColWidths = CRM_CONTACT_TABLE_COL_WIDTHS
            NewBtnTxt = "New Contact"
            LeadBtnVisible = True
            CalImpBtnVisible = True
            
        Case enScrCRMLender
        
            HeaderText = "CRM - Lenders"
            TableHeadingText = CRM_LENDER_TABLE_TITLES
            TableColWidths = CRM_LENDER_TABLE_COL_WIDTHS
            NewBtnTxt = "New Lender"
            
       Case enScrCRMProject
       
            HeaderText = "CRM - Projects"
            TableHeadingText = CRM_PROJECT_TABLE_TITLES
            TableColWidths = CRM_PROJECT_TABLE_COL_WIDTHS
            NewBtnTxt = "New Project"
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
            .Text = HeaderText
            .Style = HEADER_STYLE
            .Visible = True
        End With

        With .Table
            .ColWidths = TableColWidths
            .Left = GENERIC_TABLE_LEFT
            .Top = GENERIC_TABLE_TOP
            .HPad = GENERIC_TABLE_ROWOFFSET
            .VPad = GENERIC_TABLE_COLOFFSET
            .SubTableVOff = 50
            .SubTableHOff = 20
            .HeadingText = TableHeadingText
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
    
    With BtnCRMNewItem

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = CRM_BTN_MAIN_1_LEFT
        .Top = CRM_BTN_MAIN_1_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain1"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnCRMNewItem & ":0" & """)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = NewBtnTxt
    End With
    
    With BtnCRMContCalImp

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = CRM_BTN_MAIN_2_LEFT
        .Top = CRM_BTN_MAIN_2_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain2"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnCRMContCalImport & ":0" & """)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Calendly File Import"
        .Visible = CalImpBtnVisible
    End With
    
    With BtnCRMContShwLead

        .Height = GENERIC_BUTTON_HEIGHT
        .Left = CRM_BTN_MAIN_3_LEFT
        .Top = CRM_BTN_MAIN_3_TOP
        .Width = GENERIC_BUTTON_WIDTH
        .Name = "BtnMain3"
        .OnAction = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & enBtnCRMContShwLeads & ":0" & """)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Show Only Leads"
        .Visible = LeadBtnVisible
    End With
    
    ButtonFrame.Buttons.Add BtnCRMNewItem
    ButtonFrame.Buttons.Add BtnCRMContCalImp
    ButtonFrame.Buttons.Add BtnCRMContShwLead
    
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
Public Function RefreshList(ByVal ScreenPage As enScreenPage, Optional SortBy As String, Optional Filter As String) As Boolean
    Dim NoCols As Integer
    Dim NoRows As Integer
    Dim Workflows As ClsWorkflows
    Dim StrSortBy As String
    Dim RstWorkflowList As Recordset
    Dim y As Integer
    Dim x As Integer
    Dim AryStyles() As String
    Dim AryOnAction() As String
    Dim OpenItmBtn As EnumBtnNo
    Dim ItemIndex As String
    
    Const StrPROCEDURE As String = "RefreshList()"

    On Error GoTo ErrorHandler

    ModLibrary.PerfSettingsOn

    'load page specific data
    Select Case ScreenPage
        Case enScrCRMClient
            OpenItmBtn = enBtnCRMOpenItem
            ItemIndex = "ClientNo"
        Case enScrCRMSPV
            OpenItmBtn = enBtnCRMOpenItem
            ItemIndex = "SPVNo"
        Case enScrCRMContact
            OpenItmBtn = enBtnCRMOpenItem
            ItemIndex = "ContactNo"
        Case enScrCRMLender
            OpenItmBtn = enBtnCRMOpenItem
            ItemIndex = "LenderNo"
        Case enScrCRMProject
            OpenItmBtn = enBtnCRMOpenItem
            ItemIndex = "ProjectNo"
    End Select
    
    Set Workflows = New ClsWorkflows
    
    Set RstWorkflowList = GetCRMData(ScreenPage, StrSortBy, Filter)
    
    With RstWorkflowList
        If .RecordCount = 0 Then GoTo GracefulExit
        .MoveLast
        .MoveFirst
    End With
    
    With MainFrame.Table
        .RstText = RstWorkflowList
        .NoRows = RstWorkflowList.RecordCount
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add GENERIC_TABLE_HEADER
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
                AryStyles(x, y) = CRM_TABLE_STYLES
                AryOnAction(x, y) = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & OpenItmBtn & ":" & .Fields(ItemIndex) & """)'"
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
' Method GetCRMData
' Gets data for workflow list
'---------------------------------------------------------------
Private Function GetCRMData(ByVal ScreenPage As enScreenPage, StrSortBy As String, Optional Filter As String) As Recordset
    Dim AryBtn() As String
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim FilterCol As String
    Dim FilterStr As String
    Dim SQL As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String

    If Filter <> "" Then
        AryBtn = Split(Filter, ":")
        FilterCol = AryBtn(0)
        FilterStr = AryBtn(1)
    End If
    
    Select Case ScreenPage
        Case enScrCRMClient
            SQL = "SELECT ClientNo, Name, PhoneNo, Url FROM TblClient"
        Case enScrCRMSPV
            SQL = "SELECT SPVNo, Name FROM TblSPV"
        Case enScrCRMContact
            SQL = "SELECT TblContact.ContactNo, TblContact.ContactName, TblContact.ContactType, TblContact.Organisation, TblContact.Position, TblContact.Phone1 " _
                    & "FROM TblContact"
            If Filter <> "" Then
                SQL = SQL & " WHERE " & FilterCol & " = '" & FilterStr & "'"
            End If
        Case enScrCRMProject
            SQL = "SELECT TblProject.ProjectNo, TblClient.Name, TblSPV.Name, TblCBSUser.UserName " _
                    & "FROM ((TblProject LEFT JOIN TblClient ON TblProject.ClientNo = TblClient.ClientNo) LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo"
        Case enScrCRMLender
            SQL = "SELECT LenderNo, Name, PhoneNo, LenderType, Address FROM TblLender"
    
    End Select

    Set RstWorkflow = ModDatabase.SQLQuery(SQL)
    
    Set GetCRMData = RstWorkflow
    
End Function

' ===============================================================
' OpenItem
' Opens CRM item dependant on page.  New item created if index = 0
' ---------------------------------------------------------------
Public Function OpenItem(ScreenPage As enScreenPage, Optional Index As String) As Boolean
    Dim CRMItem As Object
    
    Const StrPROCEDURE As String = "OpenItem()"

    On Error GoTo ErrorHandler

    Select Case ScreenPage
    
        Case enScrCRMClient
            Set CRMItem = New ClsClient
        Case enScrCRMSPV
            Set CRMItem = New ClsSPV
        Case enScrCRMContact
            Set CRMItem = New ClsContact
       Case enScrCRMLender
            Set CRMItem = New ClsLender
        Case enScrCRMProject
            Set CRMItem = New ClsProject
    End Select

    With CRMItem
        If Index <> "" Then
            .DBGet CInt(Index)
            .DisplayForm
        Else
            .DBNew
        End If
    End With
    
    OpenItem = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    OpenItem = False

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
' CalendlyImport
' Imports Calendly CSV File
' ---------------------------------------------------------------
Public Function CalendlyImport() As Boolean
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim NoRows As Integer
    Dim CalendlyFile As Workbook
    Dim calendlySht As Worksheet
    Dim ColFirstName As Integer
    Dim ColLastName As Integer
    Dim ColEmail As Integer
    Dim ColEventType As Integer
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CalendlyImport()"

    On Error GoTo ErrorHandler

    Set Fldr = Application.FileDialog(msoFileDialogFilePicker)
    With Fldr
        .Title = "Select the Calendly Import File"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv", 1
        .AllowMultiSelect = False
        .ButtonName = "Import"
        .InitialFileName = Application.DefaultFilePath
        
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With

    Set CalendlyFile = Workbooks.Open(FilePath)
    
    Debug.Assert Not CalendlyFile Is Nothing
    
    If CalendlyFile Is Nothing Then Err.Raise HANDLED_ERROR, , "No Calendly file not found"
    
    Set calendlySht = CalendlyFile.Worksheets(1)
    
    Debug.Assert Not calendlySht Is Nothing
    
    If CalendlyFile Is Nothing Then Err.Raise HANDLED_ERROR, , "No Calendly sheet not found"
    
    With calendlySht
        NoRows = .UsedRange.Rows.Count - 1
        
        Debug.Assert NoRows > 0
        
        ColFirstName = .Range("1:1").Find("Invitee First Name").Column
        ColLastName = .Range("1:1").Find("Invitee Last Name").Column
        ColEmail = .Range("1:1").Find("Invitee Email").Column
        ColEventType = .Range("1:1").Find("Event Type Name").Column
        
        Debug.Assert ColFirstName > 0 And ColLastName > 0 And ColEmail > 0 And ColEventType > 0
        
        For i = 2 To NoRows + 1
            if .Range.Cells(i,coleventype
        
        
        
        Next
    End With
    
    CalendlyImport = True
    
    Set CalendlyFile = Nothing
    Set calendlySht = Nothing

Exit Function

ErrorExit:

    '***CleanUpCode***
    CalendlyImport = False

    Set CalendlyFile = Nothing
    Set calendlySht = Nothing
    
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
