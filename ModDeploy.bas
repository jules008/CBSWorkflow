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
    If DB Is Nothing Then
    Set DB = OpenDatabase(GetDocLocalPath(ThisWorkbook.Path) & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    End If
    
    'undo
    DB.Execute "ALTER TABLE TblProject DROP COLUMN FirstClientInt "
    DB.Execute "ALTER TABLE TblProject DROP COLUMN SecondClientRef "
    DB.Execute "ALTER TABLE TblProject DROP COLUMN Facilitator "
    DB.Execute "ALTER TABLE TblProject ADD COLUMN CBSComPC integer"
    DB.Execute "UPDATE TblProject SET CBSComPC = CBSCommission"
    DB.Execute "ALTER TABLE TblProject DROP COLUMN CBSCommission"
    
    
    Stop
    'update
    DB.Execute "ALTER TABLE TblProject ADD COLUMN FirstClientInt integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN SecondClientRef integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN Facilitator integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN CBSCommission integer"
    DB.Execute "UPDATE TblProject SET CBSCommission = CBSComPC"
    DB.Execute "ALTER TABLE TblProject DROP COLUMN CBSComPC"
    
    DB.Execute "DELETE * FROM TblCBSUser"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('1', 'Jason Way')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('2', 'Heather Critchlow')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('3', ' Steven Dunn')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('4', 'Hari Patel')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('5', 'Emma Flindell')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('6', 'Matt Harrison')"
    
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
            
    
    ' ========================================================================================
    ' Database commands
    ' ----------------------------------------------------------------------------------------
    DB.Execute "ALTER TABLE TblProject ADD COLUMN FirstClientInt integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN SecondClientRef integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN Facilitator integer"
    DB.Execute "ALTER TABLE TblProject ADD COLUMN CBSCommission integer"
    DB.Execute "UPDATE TblProject SET CBSCommission = CBSComPC"
    DB.Execute "ALTER TABLE TblProject DROP COLUMN CBSComPC"
    
    DB.Execute "DELETE * FROM TblCBSUser"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('1', 'Jason Way')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('2', 'Heather Critchlow')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('3', ' Steven Dunn')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('4', 'Hari Patel')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('5', 'Emma Flindell')"
    DB.Execute "UPDATE TblCBSUser INSERT INTO (CBSUserNo, UserName) VALUES ('6', 'Matt Harrison')"
    ' ========================================================================================
        
    'update DB Version
    With RstTable
        .Edit
        !VERSION = DB_VER
        !UpdateDB = False
        .Update
    End With
    
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
    ' ========================================================================================
    ' Database commands
    ' ----------------------------------------------------------------------------------------
    DB.Execute "ALTER TABLE TblProject DROP COLUMN FirstClientInt "
    DB.Execute "ALTER TABLE TblProject DROP COLUMN SecondClientRef "
    DB.Execute "ALTER TABLE TblProject DROP COLUMN Facilitator "
    DB.Execute "ALTER TABLE TblProject ADD COLUMN CBSComPC integer"
    DB.Execute "UPDATE TblProject SET CBSComPC = CBSCommission"
    DB.Execute "ALTER TABLE TblProject DROP COLUMN CBSCommission"
    ' ========================================================================================
    
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
Public Sub UpdateTable()
    Dim RstTable As Recordset
    Dim Fld As Field
    Dim i As Integer
    Dim x As Integer
    Dim Val As String
    Dim RngFields As Range
    Dim RngCol As Range
    
    If DB Is Nothing Then DBConnect
    
    DB.Execute "DELETE * FROM TblStepTemplate"
        
    Set RstTable = ModDatabase.SQLQuery("TblStepTemplate")
    Set RngFields = ShtTableImport.Range("A1:Y20")
    
    With RstTable
        x = 2
        Do While ShtTableImport.Cells(x, 1) <> ""
            i = 1
            .AddNew
            For Each Fld In RstTable.Fields
                Set RngCol = RngFields.Find(CStr(Fld.Name), , , xlWhole, xlByRows, xlNext, False)
                
                If RngCol Is Nothing Then
                    Debug.Print Fld.Name & " not found"
                Else
                    Val = ShtTableImport.Cells(x, RngCol.Column)
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
    ShtTableImport.Visible = xlSheetHidden
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

' ===============================================================
' UpdateTableData
' Updates the table data without changing version number of DB
' ---------------------------------------------------------------
Public Sub UpdateTableData()
    Dim RstProject As Recordset
    Dim RstWorkflow As Recordset
    Dim NoComplete As Integer
    Dim NoSteps As Integer
    Dim AryStepNo() As String
    Dim SumProgress As Single
    Dim CntProgress As Integer
    
    Set RstProject = ModDatabase.SQLQuery("TblProject")
    
    If DB Is Nothing Then DBConnect
    If Not DEV_MODE Or ShtSettings.ChkDebugOride Then
    
        Do While Not RstProject.EOF
            SumProgress = 0
            CntProgress = 0
            Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE ProjectNo = " & RstProject!ProjectNo)
            
            With RstWorkflow
                Do While Not .EOF
                    .Edit
                    If Not IsNull(!CurrentStep) And !CurrentStep <> "" Then
                        AryStepNo = Split(!CurrentStep, ".")
                        
                        Select Case AryStepNo(0)
                            Case 1
                                NoSteps = 20
                            Case 2
                                NoSteps = 91
                            Case 3
                                NoSteps = 16
                            Case 4
                                NoSteps = 16
                            Case 5
                                NoSteps = 15
                            Case 6
                                NoSteps = 15
                        End Select
                        
                        NoComplete = AryStepNo(1)
                        
                        If NoSteps > 0 Then
                            !Progress = NoComplete / NoSteps * 100
                            SumProgress = SumProgress + !Progress
                            CntProgress = CntProgress + 1
                        End If
                        .Update
                    End If
                    .MoveNext
                Loop
                DB.Execute "UPDATE TblWorkflow SET Progress = " & SumProgress / CntProgress & " WHERE ProjectNo = " & RstProject!ProjectNo & " And WorkflowType = 'enProject'"
            End With
            RstProject.MoveNext
        Loop
            
        DB.Execute "DELETE * FROM TblWorkflow WHERE ProjectNo = 0"
        DB.Execute "DELETE * FROM TblStepTemplate"
        DB.Execute "UPDATE TblWorkflow SET Progress = 0 WHERE Progress IS NULL"
        
        ModDeploy.UpdateTable
        
        If DEV_MODE Then
            ShtSettings.ChkDebugOride = False
        Else
            ShtSettings.ChkUpdateDB = False
            Application.EnableEvents = False
            ActiveWorkbook.Save
            Application.EnableEvents = True
        End If
    End If
    
    Set RstProject = Nothing
    Set RstWorkflow = Nothing
End Sub
