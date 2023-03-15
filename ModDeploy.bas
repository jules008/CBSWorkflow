Attribute VB_Name = "ModDeploy"
'===============================================================
' Module ModDeploy
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 05 Nov 20
'===============================================================

Option Explicit
Dim Tables() As String
Dim OldTables() As String

Private Const StrMODULE As String = "ModDeploy"

Public Sub QueryTest()
    ModStartUp.ReadINIFile
    If DB Is Nothing Then
    Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    End If
    
    UndoScript
    Stop
    UpdateScript
End Sub

Public Sub UpdateScript()
    Dim SQL1 As String
    Dim SQL2 As String
    Dim SQL3 As String
    Dim SQL4 As String
    Dim SQL5 As String
    Dim Query1 As QueryDef
    Dim Query2 As QueryDef
    Dim Query3 As QueryDef
    Dim Query4 As QueryDef
    Dim Query5 As QueryDef

    SQL1 = "Select " _
        & "    TblCBSUser.UserName As CaseManager, " _
        & "    Count(TblProject.ProjectNo) As NoCases " _
        & "From " _
        & "    TblProject Inner Join " _
        & "    TblCBSUser On TblCBSUser.CBSUserNo = TblProject.CaseManager " _
        & "Where " _
        & "    (TblProject.CompleteDate = 0 Or " _
        & "        TblProject.CompleteDate Is Null) " _
        & "Group By " _
        & "    TblCBSUser.UserName "
        
    SQL2 = "Select " _
        & "    TblCBSUser.UserName As ClientIntroducer, " _
        & "    Count(TblProject.ProjectNo) As NoCases " _
        & "From " _
        & "    TblProject Inner Join " _
        & "    TblCBSUser On TblCBSUser.CBSUserNo = TblProject.FirstClientInt " _
        & "Where " _
        & "    (TblProject.CompleteDate = 0 Or " _
        & "        TblProject.CompleteDate Is Null) " _
        & "Group By " _
        & "    TblCBSUser.UserName "

    SQL3 = "Select " _
        & "    Count (Count_ProjectNo)  As Active " _
        & "From " _
        & "    (Select Distinct " _
        & "         TblWorkflow.ProjectNo As [Count_ProjectNo] " _
        & "     From " _
        & "         TblWorkflow " _
        & "     Where " _
        & "         TblWorkflow.Status <> 'enComplete' And " _
        & "         TblWorkflow.ProjectNo <> 0) "

    SQL4 = "Select " _
        & "    Count(ClosedWeek.UProjectNo) as Closed " _
        & "From " _
        & "    (Select Distinct " _
        & "         TblProject.ProjectNo As UProjectNo, " _
        & "         TblProject.CompleteDate " _
        & "     From " _
        & "         TblWorkflow Right Join " _
        & "         TblProject On TblWorkflow.ProjectNo = TblProject.ProjectNo " _
        & "     Where " _
        & "         DatePart('ww', TblProject.CompleteDate) = DatePart('ww', Now())) As ClosedWeek "

    SQL5 = "Select " _
        & "    TblWorkflow.LoanType As [Avg_LoanType], " _
        & "    Avg(DateDiff('d', TblProject.StartDate, TblProject.CompleteDate)) As NoDays " _
        & "From " _
        & "    TblProject Right Join " _
        & "    TblWorkflow On TblWorkflow.ProjectNo = TblProject.ProjectNo " _
        & "Where " _
        & "    TblProject.ProjectNo > 0 And " _
        & "    TblWorkflow.LoanType Is Not Null And " _
        & "    TblProject.CompleteDate <> 0 " _
        & "Group By " _
        & "    TblWorkflow.LoanType "

    Set Query1 = New QueryDef
    Set Query2 = New QueryDef
    Set Query3 = New QueryDef
    Set Query4 = New QueryDef
    Set Query5 = New QueryDef
    
    With Query1
        .SQL = SQL1
        .Name = "CM Cases"
    End With
    
    With Query2
        .SQL = SQL2
        .Name = "CI Cases"
    End With
    
    With Query3
        .SQL = SQL1
        .Name = "Active"
    End With
    
    With Query4
        .SQL = SQL2
        .Name = "Closed"
    End With
    
    With Query5
        .SQL = SQL3
        .Name = "ProjTimeAve"
    End With

    DB.QueryDefs.Append Query1
    DB.QueryDefs.Append Query2
    DB.QueryDefs.Append Query3
    DB.QueryDefs.Append Query4
    DB.QueryDefs.Append Query5

    DB.Execute "CREATE TABLE TblTrendData"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN DataDate Date"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN Open Integer"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN Closed Integer"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN AveDev Integer"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN AveBridge Integer"
    DB.Execute "ALTER TABLE TblTrendData ADD COLUMN AveComm Integer"
    
End Sub

Public Sub UndoScript()
	
    If DEV_MODE Then
        On Error Resume Next
    Else
        On Error GoTo ErrorHandler
    End If
    
    DB.QueryDefs.Delete "CM Cases"
    DB.QueryDefs.Delete "CI Cases"
    DB.QueryDefs.Delete "Active"
    DB.QueryDefs.Delete "Closed"
    DB.QueryDefs.Delete "ProjTimeAve"

    DB.Execute "DROP TABLE TblTrendData"
Exit Sub

ErrorHandler:
    Stop
End Sub

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Function UpdateDBScript() As Boolean
    Dim FSO As FileSystemObject
    Dim Message As String
    Dim RstUpdate As Recordset
    Dim i As Integer
    
    Set FSO = New FileSystemObject
    Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    
    Const StrPROCEDURE As String = "UpdateDBScript()"
    
    Dim RstTable As Recordset
    Dim Fld As Field

    On Error GoTo ErrorExit
    
    Err.Clear
    
    If Not UpdateDBScriptUndo Then Err.Raise HANDLED_ERROR
    
    If DB Is Nothing Then
        Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    End If
    
    Set RstTable = DB.OpenRecordset("TblDBVersion", dbOpenDynaset)
    
    If DateDiff("d", RstTable!LastBackUp, Now()) <> 0 Then BackupFiles
    
    'check preceding DB Version
    If RstTable!VERSION <> OLD_DB_VER Then
        MsgBox "Failed update, database needs to be Version " & OLD_DB_VER & " to continue", vbOKOnly + vbCritical
        Exit Function
    End If
            
    'update DB Version
    With RstTable
        .Edit
        !VERSION = DB_VER
        !UpdateDB = False
        .Update
    End With
    Set RstTable = Nothing
    
    UpdateScript
    
        MsgBox "Database successfully updated to Version " & DB_VER, vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    DB.Close
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScript = True
    
Exit Function

ErrorExit:
   
    Debug.Print "There was an error with the database update.  Error " & Err.Number & ", " & Err.Description, vbCritical, APP_NAME
    If Not UpdateDBScriptUndo Then Err.Raise HANDLED_ERROR
    
    If Not DEV_MODE Then MsgBox "Database changes have been reversed.  Please restore previous version of FIRES", vbCritical, APP_NAME
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScript = False
    Stop
    Resume
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
    If DEBUG_MODE Then Resume
    Else
        Resume ErrorExit
    End If
End Function
              
' ===============================================================
' UpdateDBScriptUndo
' Script to update DB
' ---------------------------------------------------------------
Public Function UpdateDBScriptUndo() As Boolean
    Const StrPROCEDURE As String = "UpdateDBScriptUndo()"
    
    Dim RstTable As Recordset
    Dim Fld As DAO.Field
    
    On Error GoTo ErrorHandler
    
    If DB Is Nothing Then
        Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    End If
    
    Set RstTable = DB.OpenRecordset("TblDBVersion", dbOpenDynaset)
    
        
    If RstTable.Fields(0) <> DB_VER Then
'        UpdateDBScriptUndo = True
'        Exit Function
    End If
    
    With RstTable
        .Edit
        !VERSION = OLD_DB_VER
        .Update
    End With
    
    On Error Resume Next
    UndoScript
    
    DB.Close
    Set RstTable = Nothing
    Set DB = Nothing
    UpdateDBScriptUndo = True

Exit Function

ErrorExit:

    MsgBox "Database changes have been reversed.  Please restore previous version of FIRES", vbCritical, APP_NAME
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScriptUndo = False
    Stop
    If DEBUG_MODE Then Resume
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' UpdateTable
' Updates entire table from table update sheet
' ---------------------------------------------------------------
Public Sub UpdateTable(StrTable As String, RngImport As String)
    Dim RstTable As Recordset
    Dim Fld As Field
    Dim i As Integer
    Dim x As Integer
    Dim Val As String
    Dim RngFields As Range
    Dim RngCol As Range
    Dim ShtImport As Worksheet
    
    If DB Is Nothing Then DBConnect
    
    DB.Execute "DELETE * FROM " & StrTable
        
    Set RstTable = ModDatabase.SQLQuery(StrTable)
    
    Set ShtImport = Worksheets(StrTable)
    Set RngFields = ShtImport.Range(RngImport)
    
    With RstTable
        x = 2
        Do While ShtImport.Cells(x, 1) <> ""
            i = 1
            .AddNew
            For Each Fld In RstTable.Fields
                Set RngCol = RngFields.Find(CStr(Fld.Name), , , xlWhole, xlByRows, xlNext, False)
                
                If RngCol Is Nothing Then
                    Debug.Print Fld.Name & " not found"
                Else
                    Val = ShtImport.Cells(x, RngCol.Column)
                    Debug.Print "Col: "; RngCol.Column, "Row: "; x, Fld.Name, Val, Fld.Type
                    
                    Select Case Fld.Type
                        Case 1
                            If Val <> "" Then Fld = CBool(Val)
                        Case 4
                            If Val <> "" Then Fld = CInt(Val)
                        Case 10, 12
                            If Val <> "" Then Fld = CStr(Val)
                        Case 8
                            If IsDate(Val) Then Fld = CDate(Val)
                       
                    End Select
    
                i = i + 1
                End If
            Next
            x = x + 1
            .Update
        Loop
    End With
    ShtImport.Visible = xlSheetHidden
    Set ShtImport = Nothing
End Sub

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    Dim UpdateStatus As Boolean
    
    Const StrPROCEDURE As String = "GetDBVer()"

    Set DBVer = DB.OpenRecordset("TblDBVersion", dbOpenDynaset)

    GetDBVer = DBVer.Fields(0)
    Debug.Print DBVer.Fields(0)
    
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

End Function

' ===============================================================
' CopyFiles
' Copies and backs up FIRES file before update
' ---------------------------------------------------------------
Public Function CopyFiles() As Boolean
    Const StrPROCEDURE As String = "CopyFiles()"
    Dim FSO As FileSystemObject
    
    On Error GoTo ErrorHandler
    
    Set FSO = New Scripting.FileSystemObject
    
    With FSO
        If .FileExists(GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm") Then
            .CopyFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm", GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_BAK.xlsm"
            .DeleteFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm"
            ThisWorkbook.SaveAs GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm"
        Else
            Err.Raise 2022, , "FIRES file not found"
        End If
    
        .DeleteFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_NEW.xlsm"
    End With
    
    Set FSO = Nothing
    CopyFiles = True
Exit Function

ErrorExit:

    MsgBox "There was an error with the FIRES file update.", vbCritical, APP_NAME
    If Not FileCopyUndo Then Err.Raise HANDLED_ERROR
    
    Set FSO = Nothing
    CopyFiles = False
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' FileCopyUndo
' reverses file copy after failed update
' ---------------------------------------------------------------
Public Function FileCopyUndo() As Boolean
    Const StrPROCEDURE As String = "FileCopyUndo()"
    Dim FSO As FileSystemObject
    
    On Error GoTo ErrorHandler
    
    Set FSO = New Scripting.FileSystemObject
    
    With FSO
        If .FileExists(GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_BAK.xlsm") Then
            If .FileExists(GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm") Then
                If ThisWorkbook.Name = "FIRES.xlsm" Then
                    ThisWorkbook.SaveAs GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_NEW.xlsm"
                End If
                .DeleteFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm"
            End If
            .CopyFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_BAK.xlsm", GetDocLocalPath(ThisWorkbook.Path) & "\FIRES.xlsm"
            .DeleteFile GetDocLocalPath(ThisWorkbook.Path) & "\FIRES_BAK.xlsm"
                    
        End If
    End With
    
    Set FSO = Nothing
    FileCopyUndo = True
Exit Function

ErrorExit:

    MsgBox "There was an error with the FIRES file restore. Please contact Support.", vbCritical, APP_NAME
    
    Set FSO = Nothing
    FileCopyUndo = False
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

