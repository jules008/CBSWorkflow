Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef
'Public AccessApp As Access.Application

' ===============================================================
' SQLQuery
' Queries database with given SQL script
' ---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "SQLQuery()"

    On Error GoTo ErrorHandler
      
Restart:
    On Error Resume Next
    Application.StatusBar = ""

    On Error GoTo ErrorHandler
    
    If DB Is Nothing Then DBConnect
        If FaultCount1008 > 0 Then FaultCount1008 = 0
    
        Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
        Set SQLQuery = RstResults
    
    Set RstResults = Nothing
    
Exit Function

ErrorExit:

    Set RstResults = Nothing

    Set SQLQuery = Nothing
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        If CustomErrorHandler(Err.Number) Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            Resume Restart
        Else
            Err.Raise HANDLED_ERROR
        End If
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DBConnect
' Provides path to database
' ---------------------------------------------------------------
Public Function DBConnect() As Boolean
    Const StrPROCEDURE As String = "DBConnect()"

    On Error GoTo ErrorHandler
        
    If Not ReadINIFile Then Error.Raise HANDLED_ERROR
    Debug.Print Now & " - Connect to DB: " & DB_PATH & DB_FILE_NAME
    
    Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & "/" & DB_PATH & DB_FILE_NAME)
  
    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

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
' DBTerminate
' Disconnects and closes down DB connection
' ---------------------------------------------------------------
Public Function DBTerminate() As Boolean
    Const StrPROCEDURE As String = "DBTerminate()"

    On Error GoTo ErrorHandler

    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing
    
    DBTerminate = True

Exit Function

ErrorExit:

    DBTerminate = False

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
' SelectDB
' Selects DB to connect to
' ---------------------------------------------------------------
Public Function SelectDB() As Boolean
    Const StrPROCEDURE As String = "SelectDB()"

    On Error GoTo ErrorHandler
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly + vbExclamation, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    DB_PATH = FileLoc
    
    Set DlgOpen = Nothing

    SelectDB = True

Exit Function

ErrorExit:

    Set DlgOpen = Nothing
    SelectDB = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    Dim UpdateStatus As Boolean
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetDBVer = DBVer!VERSION
    
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetBackUpDate
' Returns the back up date
' ---------------------------------------------------------------
Public Function GetBackUpDate() As Date
    Dim DBVer As Recordset
    Dim UpdateStatus As Boolean
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetBackUpDate = DBVer!LastBackUp
    
    Set DBVer = Nothing
    
Exit Function

ErrorExit:

    GetBackUpDate = 0
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateSysMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub UpdateSysMsg()
    Dim RstMessage As Recordset
    
    Set RstMessage = SQLQuery("TblMessage")
    
    With RstMessage
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        .Fields("SystemMessage") = "Version " & VERSION & " - What's New" _
                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Bug Fix - Hidden Assets" _
                    & Chr(13) & ""
        
        .Fields("ReleaseNotes") = "Software Version: " & VERSION _
                    & Chr(13) & "Database Version: " & DB_VER _
                    & Chr(13) & "Date: " & VER_DATE _
                    & Chr(13) & "" _
                    & Chr(13) & "- Bug Fix - Hidden Assets - Had ANOTHER go at fixing the hidden assets bug.  Hopefully fixed now" _
                    & Chr(13) & ""
        .Update
    End With
    
    'reset read flags
    DB.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
    
    Set RstMessage = Nothing

End Sub

' ===============================================================
' ShowUsers
' Show users logged onto system
' ---------------------------------------------------------------
Public Sub ShowUsers()
    Dim RstUsers As Recordset
    
    Set RstUsers = SQLQuery("TblUsers")
    
    With RstUsers
        Do While Not .EOF
            .MoveNext
        Loop
    End With
    
    Set RstUsers = Nothing
End Sub

' ===============================================================
' CreateStatsTbl
' updates the statistics table in the database
' ---------------------------------------------------------------
Public Function CreateStatsTbl(AryOutput() As Variant, AryMembers() As Variant, AryTotals() As Variant) As Boolean
    Dim RstRepData As Recordset
    Dim Lb As Integer
    Dim Ub As Integer
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CreateStatsTbl()"

    On Error GoTo ErrorHandler

    Lb = LBound(AryMembers, 2)
    Ub = UBound(AryMembers, 2)
    
    DB.Execute "DELETE * FROM TblRepData"
    
    Set RstRepData = ModDatabase.SQLQuery("TblRepData")
    
    With RstRepData
        For i = Lb To Ub
            .AddNew
            !StudentID = AryMembers(5, i)
            !Watch = AryMembers(4, i)
            If AryMembers(6, i) = "Active" Then !Active = True Else !Active = False
            !Position = AryMembers(2, i)
            !QualsNeeded = AryTotals(i, 0)
            !ReqQualsGnd = AryTotals(i, 1)
            !ExtraQuals = AryTotals(i, 2)
            !PCQuald = AryTotals(i, 3)
            If !QualsNeeded = !ReqQualsGnd Then !QIP = True Else !QIP = False
            .Update
        Next
    End With

    CreateStatsTbl = True

    Set RstRepData = Nothing

Exit Function

ErrorExit:

    Set RstRepData = Nothing
    '***CleanUpCode***
    CreateStatsTbl = False

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
' BackupFiles
' Backs up database at predetermined intervals
' ---------------------------------------------------------------
Public Function BackupFiles() As Boolean
    Dim FSO As FileSystemObject
    
    Const StrPROCEDURE As String = "BackupFiles()"

    On Error GoTo ErrorHandler

    Set FSO = New FileSystemObject
            
    FSO.CopyFile GetDocLocalPath(ThisWorkbook.Path) & "\System Files\" & DB_FILE_NAME & ".accdb", GetDocLocalPath(ThisWorkbook.Path) & BAK_FILES & DB_FILE_NAME & " BAK-" & Format(Now, "yy-mm-dd hhmm") & ".accdb", True
            
    Set FSO = Nothing



    BackupFiles = True


Exit Function

ErrorExit:

    Set FSO = Nothing
    BackupFiles = True

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ifTableExists
' Check if database table exists
' ---------------------------------------------------------------
Public Function ifTableExists(tblName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    If DB Is Nothing Then DBConnect
    
    ifTableExists = IsObject(DB.TableDefs(tblName))
Exit Function

ErrorHandler:
    ifTableExists = False
End Function
