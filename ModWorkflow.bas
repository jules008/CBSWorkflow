Attribute VB_Name = "ModWorkflow"
'===============================================================
' Module ModWorkflow
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Jun 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModWorkflow"

' ===============================================================
' KeyWords
' Replaces <keywords> with defined data
' ---------------------------------------------------------------
Private Function KeyWords(Keyword As String, ByRef Workflow As ClsWorkflow) As String
    Dim x As String
    
    Const StrPROCEDURE As String = "KeyWords()"

    On Error GoTo ErrorHandler

    Select Case Keyword
        Case "{ClientEmail}"
            KeyWords = Workflow.Parent.Client.Contacts.PrimaryContact.EmailAddress
            
        Case "{CurrentUserName}"
            KeyWords = CurrentUser.UserName
        
        Case "{CurrentUserRole}"
            KeyWords = CurrentUser.Position
        
        Case "{CurrentUserPhone}"
            KeyWords = CurrentUser.PhoneNo
        
        Case "{ClientName}"
            KeyWords = Workflow.Parent.Client.Contacts.PrimaryContact.ContactName

        Case "{BusinessHead}"
            KeyWords = "eflindell@cbscapital.co.uk"

        Case "{Director}"
            KeyWords = "sdunn@cbscapital.co.uk"

    End Select

GracefulExit:

Exit Function

ErrorExit:
    KeyWords = "Error"

Exit Function

ErrorHandler:
    If KeyWords = "" Then
        Err.Description = "Keyword not found - " & Keyword
        MsgBox "WARNING: keyword " & Keyword & " has not being found.  Check in the Administration " _
        & "Section and ensure that it has been assigned correctly.", vbExclamation, APP_NAME
        Resume GracefulExit
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' EmailLookUp
' Replaces {EmailLookUp} with defined data
' ---------------------------------------------------------------
Private Function EmailLookUp(EmailAddress As String) As String
    Dim RstEmail As Recordset
    Dim OutputText As String
    
    Const StrPROCEDURE As String = "EmailLookUp()"

    On Error GoTo ErrorHandler
        
    Set RstEmail = ModDatabase.SQLQuery("SELECT EmailAddress FROM TblEmailAddress WHERE EmailName = '" & EmailAddress & "'")
    
    With RstEmail
        If .RecordCount > 0 Then EmailLookUp = !EmailAddress
        OutputText = !EmailAddress
    End With
    EmailLookUp = OutputText
    
    Set RstEmail = Nothing
    
Exit Function

ErrorExit:
    EmailLookUp = "Error"
    Set RstEmail = Nothing

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
' ReplaceKeyWords
' parses text and replaces keywords with specified data
' ---------------------------------------------------------------
Public Function ReplaceKeyWords(InputText As String, ByRef Workflow As ClsWorkflow) As String
    Dim Keyword As String
    Dim KeyStart As Integer
    Dim KeyEnd As Integer
    Dim DataItem As String
    Dim OutputText As String
    
    Const StrPROCEDURE As String = "ReplaceKeyWords()"

    On Error GoTo ErrorHandler
    
    OutputText = InputText
    Do While InStr(1, OutputText, "{", vbTextCompare)
        KeyStart = InStr(1, OutputText, "{", vbTextCompare)
        KeyEnd = InStr(1, OutputText, "}", vbTextCompare)
        Keyword = Mid(OutputText, KeyStart, KeyEnd - KeyStart + 1)
        DataItem = KeyWords(Keyword, Workflow)
        OutputText = Replace(OutputText, Keyword, DataItem)
    Loop
    
    Do While InStr(1, OutputText, "{", vbTextCompare)
        KeyStart = InStr(1, OutputText, "{", vbTextCompare)
        KeyEnd = InStr(1, OutputText, "}", vbTextCompare)
        Keyword = Mid(OutputText, KeyStart, KeyEnd - KeyStart + 1)
        DataItem = EmailLookUp(Keyword)
        OutputText = Replace(OutputText, Keyword, DataItem)
        
    Loop
    
    ReplaceKeyWords = OutputText

Exit Function

ErrorExit:

    '***CleanUpCode***
    ReplaceKeyWords = "Error"

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
' ProcessDataInput
' Takes data input during workflow steps and saves them
' ---------------------------------------------------------------
Public Function ProcessDataInput() As Boolean
    Dim DataInput As String
    Dim DataDest() As String
    Dim IndexNo As String
    Dim IndexVal As String
    Dim DataFormat As String
    Dim Table As String
    Dim TblField As String
    Dim SQL As String
    Dim Step As ClsStep
    Dim obj As Object
    
    Const StrPROCEDURE As String = "ProcessDataInput()"

    On Error GoTo ErrorHandler
    
    Set Step = ActiveWorkFlow.ActiveStep
    
    With Step
        .Parent.DBSave
        DataInput = .DataItem
        DataDest = Split(ReplaceKeyWords(.DataDest, .Parent), ".")
        Table = DataDest(0)
        TblField = DataDest(1)

        Select Case Table
            Case "Workflow"
                IndexNo = "WorkflowNo"
                IndexVal = .Parent.WorkflowNo
        End Select
        
        SQL = "UPDATE Tbl" & Table & " SET " & TblField & " = '" & DataInput & "' WHERE " & IndexNo & " = " & IndexVal
        DB.Execute SQL
        
        'Debug.Print SQL
        
        Select Case Table
            Case "Workflow"
                .Parent.DBGet (.Parent.WorkflowNo)
        End Select
    End With
    
    ProcessDataInput = True
    
    Set Step = Nothing
    Set obj = Nothing
Exit Function

ErrorExit:

    Set Step = Nothing
    Set obj = Nothing
    
    ProcessDataInput = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, CustomData:=DataInput) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

