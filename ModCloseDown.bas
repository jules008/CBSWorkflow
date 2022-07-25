Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
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

Private Const StrMODULE As String = "ModCloseDown"

' ===============================================================
' Terminate
' Closedown processing
' ---------------------------------------------------------------
Public Function Terminate() As Boolean
    Dim Frame As ClsUIFrame
'    Dim DashObj As ClsUIDashObj
    Dim MenuItem As ClsUIMenuItem
    Dim Lineitem As ClsUILineitem

    Const StrPROCEDURE As String = "Terminate()"

    On Error Resume Next

    ShtMain.Unprotect PROTECT_KEY

'    CurrentUser.LogUserOff
    SYSTEM_CLOSING = True

    If Not EndGlobalClasses Then Err.Raise HANDLED_ERROR

    Application.DisplayFullScreen = False

    Set MainScreen = Nothing

'    If Not CurrentUser Is Nothing Then Set CurrentUser = Nothing

    ModDatabase.DBTerminate
    DeleteAllShapes

    Terminate = True

Exit Function

ErrorExit:

    ModDatabase.DBTerminate
    DeleteAllShapes
    Application.DisplayFullScreen = False
    
    Terminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DeleteAllShapes
' Deletes all shapes on screen except templates
' ---------------------------------------------------------------
Private Sub DeleteAllShapes()
    Dim i As Integer
    
    Const StrPROCEDURE As String = "DeleteAllShapes()"

    On Error Resume Next

    Dim Shp As Shape
    
    For i = ShtMain.Shapes.Count To 1 Step -1
    
        Set Shp = ShtMain.Shapes(i)
        
        If Left(Shp.Name, 8) <> "TEMPLATE" Then Shp.Delete
    Next

End Sub

' ===============================================================
' EndGlobalClasses
' Terminates all global classes
' ---------------------------------------------------------------
Private Function EndGlobalClasses() As Boolean
    Const StrPROCEDURE As String = "EndGlobalClasses()"

    On Error GoTo ErrorHandler

    
    EndGlobalClasses = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    EndGlobalClasses = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

