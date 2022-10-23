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
'        Case "<MemberName>"
'            KeyWords = Workflow.Member.DisplayName
'
'        Case "<MemberFirstName>"
'            KeyWords = Workflow.Member.FirstName
'
'        Case "<MemberLastName>"
'            KeyWords = Workflow.Member.LastName
'
'        Case "<MemberEmail>"
'            KeyWords = Workflow.Member.Email
'
'        Case "<MemberSSN>"
'            KeyWords = Replace(Workflow.Member.SSN, "-", "")
'
'        Case "<MemberSSN->"
'            KeyWords = Workflow.Member.SSN
'
'        Case "<MemberShortSSN>"
'            KeyWords = Format(Right(Workflow.Member.SSN, 4), "0000")
'
'        Case "<WatchShort>"
'            KeyWords = Left(Workflow.Member.Watch, 1)
'
'        Case "<StudentID>"
'            KeyWords = Workflow.Member.StudentID
'
'        Case "<Grade>"
'            KeyWords = Workflow.Member.Grade
'
'        Case "<Rank>"
'            With Workflow.Member
'                If .Grade = "S06" Or .Grade = "S07" Then KeyWords = "FF"
'                If .Grade = "S08" Then KeyWords = "CM"
'                If .Grade = "S09" Then KeyWords = "SC"
'                If .Grade = "S10" Then KeyWords = "AC"
'            End With
'
'        Case "<CDCEnrolNo>"
'            KeyWords = Format(Workflow.CDC.CDCEnrolNo)
'
'        Case "<CDCExamDate>"
'            KeyWords = Format(Workflow.CDC.ExamDate, "dd mmm yy")
'
'        Case "<CDCExamTime>"
'            KeyWords = Format(Workflow.CDC.ExamTime, "hh:mm")
'
'        Case "<CDCEnrolDate>"
'            KeyWords = Format(Workflow.CDC.EnrolDate, "dd mmm yy")
'
'        Case "<CDCComplDate>"
'            KeyWords = Format(Workflow.CDC.ComplDate, "dd mmm yy")
'
'        Case "<CDCStartDate>"
'            Workflow.CDC.StartDate = Now
'            Workflow.CDC.DBSave
'
'            KeyWords = Format(Workflow.CDC.StartDate, "dd mmm yy")
'
'        Case "<Date>"
'            KeyWords = Format(Now, "dd mmm yy")
'
'        Case "<SCAEmail>"
'            With Members
'                If Workflow.Member.Watch = "White" Then KeyWords = Members.FindItem(.SCWhiteA).Email
'                If Workflow.Member.Watch = "Red" Then KeyWords = Members.FindItem(.SCRedA).Email
'                If Workflow.Member.Watch = "Blue" Then KeyWords = Members.FindItem(.SCBlueA).Email
'                If Workflow.Member.Watch = "Green" Then KeyWords = Members.FindItem(.SCGreenA).Email
'            End With
'
'        Case "<SCBEmail>"
'            With Members
'                If Workflow.Member.Watch = "White" Then KeyWords = Members.FindItem(.SCWhiteB).Email
'                If Workflow.Member.Watch = "Red" Then KeyWords = Members.FindItem(.SCRedB).Email
'                If Workflow.Member.Watch = "Blue" Then KeyWords = Members.FindItem(.SCBlueB).Email
'                If Workflow.Member.Watch = "Green" Then KeyWords = Members.FindItem(.SCGreenB).Email
'            End With
'
'        Case "<ACEmail>"
'            With Members
'                If Workflow.Member.Watch = "White" Then KeyWords = Members.FindItem(.ACWhite).Email
'                If Workflow.Member.Watch = "Red" Then KeyWords = Members.FindItem(.ACRed).Email
'                If Workflow.Member.Watch = "Blue" Then KeyWords = Members.FindItem(.ACBlue).Email
'                If Workflow.Member.Watch = "Green" Then KeyWords = Members.FindItem(.ACGreen).Email
'            End With
'
'        Case "<ACTrainingEmail>"
'            With Members
'                KeyWords = Members.FindItem(.ACTraining).Email
'            End With
'
'        Case "<FChiefEmail>"
'            With Members
'                KeyWords = Members.FindItem(.FChief).Email
'            End With
'
'        Case "<ACLastName>"
'            With Members
'                If Workflow.Member.Watch = "White" Then KeyWords = Members.FindItem(.ACWhite).LastName
'                If Workflow.Member.Watch = "Red" Then KeyWords = Members.FindItem(.ACRed).LastName
'                If Workflow.Member.Watch = "Blue" Then KeyWords = Members.FindItem(.ACBlue).LastName
'                If Workflow.Member.Watch = "Green" Then KeyWords = Members.FindItem(.ACGreen).LastName
'            End With
'
'        Case "<PerfTestRef>"
'            KeyWords = Workflow.PerfTest.PerfTestRef
'
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
    Do While InStr(1, OutputText, "<", vbTextCompare)
        KeyStart = InStr(1, OutputText, "<", vbTextCompare)
        KeyEnd = InStr(1, OutputText, ">", vbTextCompare)
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

