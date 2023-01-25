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
Private ModFilter As String

' ===============================================================
' BuildScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildScreen(ByVal ScrnPage As enScreenPage, Optional Filter As String) As Boolean
    
    Const StrPROCEDURE As String = "BuildScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    ModFilter = Filter
    
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
Public Function RefreshList(ByVal ScreenPage As enScreenPage, Optional SortBy As String) As Boolean
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
    
    Set RstWorkflowList = GetCRMData(ScreenPage, SortBy)
    
    With RstWorkflowList
        If .RecordCount = 0 Then GoTo GracefulExit
        .MoveLast
        .MoveFirst
    End With
    
    With MainFrame.Table
        .RstText = RstWorkflowList
        .NoRows = RstWorkflowList.RecordCount
        .StylesColl.RemoveCollection
        .StylesColl.Add GENERIC_TABLE
        .StylesColl.Add GENERIC_TABLE_HEADER
        .RowHeight = GENERIC_TABLE_ROW_HEIGHT
        .Cells.DeleteCollection
    End With
    
    NoRows = RstWorkflowList.RecordCount
    NoCols = MainFrame.Table.NoCols
    
    ReDim AryStyles(0 To NoCols - 1, 0 To NoRows)
    ReDim AryOnAction(0 To NoCols - 1, 0 To NoRows)
    
    Debug.Assert MainFrame.Table.Cells.Count = 0
    
    With RstWorkflowList
        For x = 0 To NoCols - 1
            .MoveFirst
            For y = 0 To NoRows
                If y = 0 Then
                    AryStyles(x, y) = "GENERIC_TABLE_HEADER"
                    If x < .Fields.Count Then
                        AryOnAction(x, y) = "'ModUICRM.SortBy (""" & ScreenPage & ":" & .Fields(x).Name & """)'"
                    End If
                Else
                    AryStyles(x, y) = "GENERIC_TABLE"
                    If x < .Fields.Count Then
                AryOnAction(x, y) = "'ModUIButtonHandler.ProcessBtnClicks(""" & ScreenPage & ":" & OpenItmBtn & ":" & .Fields(ItemIndex) & """)'"
                    End If
                .MoveNext
                End If
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
Private Function GetCRMData(ByVal ScreenPage As enScreenPage, StrSortBy As String) As Recordset
    Dim AryBtn() As String
    Dim RstWorkflow As Recordset
    Dim Workflow As ClsWorkflow
    Dim FilterCol As String
    Dim FilterStr As String
    Dim SQL As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    Dim StrFilter As String

    If ModFilter <> "" Then
        AryBtn = Split(ModFilter, ":")
        FilterCol = AryBtn(0)
        FilterStr = AryBtn(1)
    End If
    
    If StrSortBy <> "" Then StrSortBy = " ORDER BY " & StrSortBy
    If ModFilter <> "" Then StrFilter = " WHERE " & FilterCol & " = '" & FilterStr & "' "
    
    Select Case ScreenPage
        Case enScrCRMClient
            SQL = "SELECT ClientNo, Name, CBS, Address, PhoneNo, Url FROM TblClient " & StrSortBy
        Case enScrCRMSPV
            SQL = "SELECT SPVNo, Name FROM TblSPV " & StrSortBy
        Case enScrCRMContact
            SQL = "SELECT TblContact.ContactNo, TblContact.ContactName, TblContact.ContactType, TblContact.Organisation, TblContact.Position, TblContact.Phone1, TblContact.EmailAddress " _
                    & "FROM TblContact " & StrFilter & StrSortBy
        Case enScrCRMProject
            SQL = "SELECT TblProject.ProjectNo, TblProject.ProjectName, TblClient.Name, TblSPV.Name, TblCBSUser.UserName " _
                    & "FROM ((TblProject LEFT JOIN TblClient ON TblProject.ClientNo = TblClient.ClientNo) LEFT JOIN TblSPV ON TblProject.SPVNo = TblSPV.SPVNo) LEFT JOIN TblCBSUser ON TblProject.CaseManager = TblCBSUser.CBSUserNo " & StrSortBy
        Case enScrCRMLender
            SQL = "SELECT LenderNo, Name, PhoneNo, LenderType, Address FROM TblLender " & StrSortBy
    
    End Select

    Set RstWorkflow = ModDatabase.SQLQuery(SQL)
    
    Set GetCRMData = RstWorkflow
    
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

    Select Case SortByData
        Case Is = "3:ClientNo"
            StrSort = "TblClient.ClientNo"
        Case Is = "3:Name"
            StrSort = "TblClient.Name"
        Case Is = "3:Url"
            StrSort = "TblClient.url"
        Case Is = "3:PhoneNo"
            StrSort = "TblClient.PhoneNo"
        Case Is = "4:SPVNo"
            StrSort = "TblSPV.SPVNo"
        Case Is = "4:Name"
            StrSort = "TblSPV.Name"
        Case Is = "5:ContactNo"
            StrSort = "TblContact.ContactNo"
        Case Is = "5:ContactName"
            StrSort = "TblContact.ContactName"
        Case Is = "5:ContactType"
            StrSort = "TblContact.ContactType"
        Case Is = "5:Organisation"
            StrSort = "TblContact.Organisation"
        Case Is = "5:Position"
            StrSort = "TblContact.Position"
        Case Is = "5:Phone1"
            StrSort = "TblContact.Phone1"
        Case Is = "5:EmailAddress"
            StrSort = "TblContact.EmailAddress"
        Case Is = "6:ProjectNo"
            StrSort = "TblProject.ProjectNo"
        Case Is = "6:ProjectName"
            StrSort = "TblProject.ProjectName"
        Case Is = "6:TblClient.Name"
            StrSort = "TblProject.CaseManager"
        Case Is = "6:TblSPV.Name"
            StrSort = "TblProject.SPVNo"
        Case Is = "6:CaseManager"
            StrSort = "TblProject.CaseManager"
        Case Is = "7:LenderNo"
            StrSort = "TblLender.LenderNo"
        Case Is = "7:Name"
            StrSort = "TblLender.Name"
        Case Is = "7:PhoneNo"
            StrSort = "TblLender.PhoneNo"
        Case Is = "7:LenderType"
            StrSort = "TblLender.LenderType"
        Case Is = "7:Address"
            StrSort = "TblLender.Address"
    End Select

    ArySort = Split(SortByData, ":")
    ScreenPage = ArySort(0)
    SortBy = ArySort(1)
    
    If Not ModUICRM.RefreshList(ScreenPage, StrSort) Then Err.Raise HANDLED_ERROR

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
    Dim FirstName As String
    Dim LastName As String
    Dim Email As String
    Dim RstContacts As Recordset
    Dim RstContMaxNo As Recordset
    Dim Import As Boolean
    
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
        
        If .Show <> -1 Then Exit Function
        FilePath = .SelectedItems(1)
    End With

    Set CalendlyFile = Workbooks.Open(FilePath)
    
    Debug.Assert Not CalendlyFile Is Nothing
    
    If CalendlyFile Is Nothing Then Err.Raise HANDLED_ERROR, , "No Calendly file not found"
    
    Set calendlySht = CalendlyFile.Worksheets(1)
    
    Debug.Assert Not calendlySht Is Nothing
    
    If CalendlyFile Is Nothing Then Err.Raise HANDLED_ERROR, , "No Calendly sheet not found"
    
    Set RstContacts = ModDatabase.SQLQuery("TblContact")
    
    With calendlySht
        NoRows = .UsedRange.Rows.Count - 1
        
        Debug.Assert NoRows > 0
        
        ColFirstName = .Range("1:1").Find("Invitee First Name").Column
        ColLastName = .Range("1:1").Find("Invitee Last Name").Column
        ColEmail = .Range("1:1").Find("Invitee Email").Column
        ColEventType = .Range("1:1").Find("Event Type Name").Column
        
        Debug.Assert ColFirstName > 0 And ColLastName > 0 And ColEmail > 0 And ColEventType > 0
'        Debug.Print NoRows
        
        For i = 2 To NoRows + 1
            Set RstContMaxNo = ModDatabase.SQLQuery("SELECT MAX(ContactNo) As MaxNo FROM TblContact")
            
            If Trim(.Cells(i, ColEventType)) = "Funding call with Heather" Then
                FirstName = Trim(.Cells(i, ColFirstName))
                LastName = Trim(.Cells(i, ColLastName))
                Email = Trim(.Cells(i, ColEmail))
                Debug.Assert FirstName <> ""
'                Debug.Print i, FirstName, LastName, Email
            End If
            
            With RstContacts
                If .RecordCount > 0 Then
                .MoveFirst
                .FindFirst ("ContactName = '" & FirstName & " " & LastName & "'")
                
                If .NoMatch Then
'                    Debug.Print "no record found"
                    Import = True
                Else
                    Dim Response As Integer
                    Response = MsgBox("A Lead is being imported with the same name as an existing contact. " _
                        & "Do you want to import this name? " & vbCrLf _
                        & "Name: " & !ContactName & vbCrLf _
                        & "Contact Type: " & !ContactType & vbCr, vbYesNo + vbInformation, "Duplicate")
                    If Response = 6 Then
                        Import = True
                    Else
                        Import = False
                        End If
                    End If
                End If
                
                If Import Then
                    Dim NextNo As Integer
                    
                    NextNo = RstContMaxNo!MaxNo + 1
                    DB.Execute "INSERT INTO TblContact (ContactNo, ContactType, ContactName, EmailAddress) " _
                                & "VALUES (" & NextNo & ", 'Lead'" & ", '" & FirstName & " " & LastName & "', '" & Email & " ') "
                    
                End If
            End With
        Next
    End With
    
    Application.DisplayAlerts = False
    CalendlyFile.Close
    Application.DisplayAlerts = True
    
    CalendlyImport = True
    Set RstContacts = Nothing
    Set CalendlyFile = Nothing
    Set calendlySht = Nothing
    Set RstContMaxNo = Nothing
Exit Function

ErrorExit:

    Application.DisplayAlerts = True
    CalendlyImport = False

    Set RstContacts = Nothing
    Set CalendlyFile = Nothing
    Set calendlySht = Nothing
    Set RstContMaxNo = Nothing
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
