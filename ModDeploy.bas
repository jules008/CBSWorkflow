Attribute VB_Name = "ModDeploy"
'===============================================================
' Module ModDeploy
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 05 Nov 20
'===============================================================

Option Explicit
Dim Tables() As String
Dim OldTables() As String

Private Const StrMODULE As String = "ModDeploy"

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Function UpdateDBScript() As Boolean
    Const StrPROCEDURE As String = "UpdateDBScript()"
    
    Dim RstTable As Recordset
    Dim fld As Field

    On Error GoTo ErrorExit
    
    Err.Clear
    
    If Not UpdateDBScriptUndo Then Err.Raise HANDLED_ERROR
    
    If DB Is Nothing Then
        Set DB = OpenDatabase(ThisWorkbook.Path & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
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
    
    ' ========================================================================================
    ' Database commands
    ' ----------------------------------------------------------------------------------------
    
    DB.Execute "DELETE * FROM TblPosition"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Fire Chief', 1)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Deputy Chief', 2)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('AC Fire Prov', 3)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('AC H&S', 3)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('AC Training', 3)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('AC Ops', 4)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Station Captain', 5)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Crew Manager', 6)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Driver Op', 7)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Firefighter', 8)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Fire Insp', 9)"
    DB.Execute "INSERT INTO TblPosition (Position, PosNo) VALUES ('Dispatch', 10)"
    
    
    ' ========================================================================================
        
        MsgBox "Database successfully updated to Version " & DB_VER, vbOKOnly + vbInformation
    
    DB.Close
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScript = True
    
Exit Function

ErrorExit:
   
    Debug.Print "There was an error with the database update.  Error " & Err.Number & ", " & Err.Description, vbCritical, APP_NAME
    If Not UpdateDBScriptUndo Then Err.Raise HANDLED_ERROR
    MsgBox "Database changes have been reversed.  Please restore previous version of FIRES", vbCritical, APP_NAME
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScript = False
    Stop
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
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
    Dim fld As DAO.Field
    
    On Error GoTo ErrorHandler
    
    If DB Is Nothing Then
        Set DB = OpenDatabase(ThisWorkbook.Path & INI_FILE_PATH & DB_FILE_NAME & ".accdb")
    End If
    
    Set RstTable = DB.OpenRecordset("TblDBVersion", dbOpenDynaset)
    
        
    If RstTable.Fields(0) <> DB_VER Then
        UpdateDBScriptUndo = True
        Exit Function
    End If
    
    With RstTable
        .Edit
        !VERSION = OLD_DB_VER
        .Update
    End With
    
    ' ========================================================================================
    ' Database commands
    ' ----------------------------------------------------------------------------------------
   
    ' ========================================================================================
    
    DB.Close
    Set RstTable = Nothing
    Set DB = Nothing
    UpdateDBScriptUndo = True

Exit Function

ErrorExit:

    MsgBox "There was an error with the database update.  Error " & Err.Number & ", " & Err.Description, vbCritical, APP_NAME
    UpdateDBScriptUndo
    MsgBox "Database changes have been reversed.  Please restore previous version of FIRES", vbCritical, APP_NAME
    
    Set DB = Nothing
    Set RstTable = Nothing
    UpdateDBScriptUndo = False
    Stop
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
        If .FileExists(ThisWorkbook.Path & "\FIRES.xlsm") Then
            .CopyFile ThisWorkbook.Path & "\FIRES.xlsm", ThisWorkbook.Path & "\FIRES_BAK.xlsm"
            .DeleteFile ThisWorkbook.Path & "\FIRES.xlsm"
            ThisWorkbook.SaveAs ThisWorkbook.Path & "\FIRES.xlsm"
        Else
            Err.Raise 2022, , "FIRES file not found"
        End If
    
        .DeleteFile ThisWorkbook.Path & "\FIRES_NEW.xlsm"
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
        If .FileExists(ThisWorkbook.Path & "\FIRES_BAK.xlsm") Then
            If .FileExists(ThisWorkbook.Path & "\FIRES.xlsm") Then
                If ThisWorkbook.Name = "FIRES.xlsm" Then
                    ThisWorkbook.SaveAs ThisWorkbook.Path & "\FIRES_NEW.xlsm"
                End If
                .DeleteFile ThisWorkbook.Path & "\FIRES.xlsm"
            End If
            .CopyFile ThisWorkbook.Path & "\FIRES_BAK.xlsm", ThisWorkbook.Path & "\FIRES.xlsm"
            .DeleteFile ThisWorkbook.Path & "\FIRES_BAK.xlsm"
                    
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

