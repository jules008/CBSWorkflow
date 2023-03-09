Attribute VB_Name = "ModLibrary"
'===============================================================
' Module ModLibrary
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Nov 22
'===============================================================

Option Explicit

 #If VBA7 Then
    Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Declare PtrSafe Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr
    Declare PtrSafe Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
 #Else
    Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
    Private Declare Function CloseClipboard Lib "user32.dll" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
 #End If

Private Const StrMODULE As String = "ModLibrary"

' ===============================================================
' ConvertHoursIntoDecimal
' Converts standard date format into decimal format
' ---------------------------------------------------------------
Public Function ConvertHoursIntoDecimal(TimeIn As Date)
    On Error Resume Next
    
    Dim TB, Result As Single
    
    TB = Split(TimeIn, ":")
    ConvertHoursIntoDecimal = TB(0) + ((TB(1) * 100) / 60) / 100
    
End Function

' ===============================================================
' EndOfMonth
' Returns the number of days in the given month
' ---------------------------------------------------------------
Function EndOfMonth(InputDate As Date) As Variant
    On Error Resume Next
    
    EndOfMonth = Day(DateSerial(Year(InputDate), Month(InputDate) + 1, 0))
End Function

' ===============================================================
' PerfSettingsOn
' turns off system functions to increase performance
' ---------------------------------------------------------------
Public Sub PerfSettingsOn()
    On Error Resume Next
    
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

' ===============================================================
' PerfSettingsOff
' turns system functions back to normal
' ---------------------------------------------------------------
Public Sub PerfSettingsOff()
    On Error Resume Next
        
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ===============================================================
' GetTextFromFile
' returns all text from a file
' ---------------------------------------------------------------
Private Function GetTextFromFile(FilePath As String) As String
    Dim FSO As Object
    Dim oFS As Object, sText As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oFS = FSO.OpenTextFile(FilePath)

    GetTextFromFile = oFS.ReadAll()

End Function

' ===============================================================
' SpellCheck
' checks spelling on forms
' ---------------------------------------------------------------
Public Sub SpellCheck(ByRef Cntrls As Collection)
    On Error Resume Next
    
    Dim RngSpell As Range
    Dim Cntrl As Control
    
    Set RngSpell = Worksheets(1).Range("A1")
    
    For Each Cntrl In Cntrls
        
        If Left(Cntrl.Name, 3) = "Txt" Then
            RngSpell = Cntrl
            RngSpell.CheckSpelling
            Cntrl = RngSpell
        End If
    Next
    Set RngSpell = Nothing
End Sub

' ===============================================================
' RecordsetPrint
' sends contents of recordset to debug window
' ---------------------------------------------------------------
Public Sub RecordsetPrint(Rst As Recordset)
    On Error Resume Next
    
    Dim DBString As String
    Dim RSTField As Field
    Dim i As Integer

    ReDim AyFields(Rst.Fields.Count)
    
    Debug.Print Rst.RecordCount
    Rst.MoveFirst
    
    For i = 0 To Rst.Fields.Count - 1
        DBString = DBString & Rst.Fields(i).Name & ":" & vbTab
    Next
    
    Debug.Print DBString
    DBString = ""
    
    Do Until Rst.EOF
        For i = 0 To Rst.Fields.Count - 1
             DBString = DBString & Rst.Fields(i).Value & ":" & vbTab
        Next
        Rst.MoveNext
        Debug.Print DBString
        DBString = ""
    Loop

End Sub

' ===============================================================
' PrintPDF
' Prints sent worksheet as a PDF
' ---------------------------------------------------------------
Public Sub PrintPDF(WSheet As Worksheet, PathAndFileName As String)
    On Error Resume Next
    
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler
    
    strFile = PathAndFileName & ".pdf"
    
    WSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, openafterpublish:=False
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler

End Sub

' ===============================================================
' CopyTextToClipboard
' Sends string to clipboard for pasting
' ---------------------------------------------------------------
Function CopyTextToClipboard(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            CopyTextToClipboard = .GetData("text")
      End Select
    End With
  End With
    
End Function

' ===============================================================
' ColourConvert
' Converts RGB colour to long
' ---------------------------------------------------------------

Public Sub ColourConvert(R As Integer, G As Integer, B As Integer)
     Dim HexColour As Long
     Dim HexColour2 As Long
     Dim Colour1 As Long
     Colour1 = RGB(R, G, B)
     HexColour = RGB(B, G, R)
     HexColour2 = RGB(R, G, B)
     
     
     Debug.Print "Hex Colour: "; Hex(HexColour)
     Debug.Print "Hex Colour2: "; Hex(HexColour2)
     Debug.Print "Dec Colour: "; Colour1

End Sub

' ===============================================================
' AddCheckBoxes
' Adds checkboxes to selected cells
' ---------------------------------------------------------------
Sub AddCheckBoxes()
    On Error Resume Next
    Dim C As Range, myRange As Range
    Set myRange = Selection
    For Each C In myRange.Cells
        ActiveSheet.CheckBoxes.Add(C.Left, C.Top, C.Width, C.Height).Select
            With Selection
                .LinkedCell = C.Address
                .Characters.Text = ""
                .Name = C.Address
            End With
            C.Select
            With Selection
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, _
                    Formula1:="=" & C.Address & "=TRUE"
                '.FormatConditions(1).Font.ColorIndex = 6 'change for other color when ticked
                '.FormatConditions(1).Interior.ColorIndex = 6 'change for other color when ticked
                '.Font.ColorIndex = 2 'cell background color = White
            End With
        Next
        myRange.Select
        Set C = Nothing
        Set myRange = Nothing
    
End Sub

' ===============================================================
' OutlookRunning
' Checks whether Outlook application is running
' ---------------------------------------------------------------
Function OutlookRunning() As Boolean
    Dim oOutlook As Object

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOutlook Is Nothing Then
        OutlookRunning = False
    Else
        OutlookRunning = True
    End If
End Function

' ===============================================================
' GetTextLineNo
' returns the number of lines in a csv or text file
' ---------------------------------------------------------------
Public Function GetTextLineNo(FileName As String) As Integer
    Dim wb As Workbook
    
    For Each wb In Workbooks
        If wb.FullName = FileName Then wb.Close (False)
    Next wb
   
    Set wb = Workbooks.Open(FileName)
    
    If Not wb Is Nothing Then
        With wb.Worksheets(1)
        
            GetTextLineNo = .Cells(.Rows.Count, "A").End(xlUp).Row
            wb.Close savechanges:=False
        End With
    End If
    
    Set wb = Nothing
End Function

' ===============================================================
' PrintDoc
' Prints any document
' ---------------------------------------------------------------
Public Function PrintDoc(FileName As String)
    Dim x As LongPtr
    
    On Error Resume Next
    
    x = ShellExecute(0, "Print", FileName, 0&, 0&, 3)

End Function

' ===============================================================
' OpenDoc
' Opens any document
' ---------------------------------------------------------------
Public Function OpenDoc(FileName As String)
    Dim x As LongPtr
    
'    On Error Resume Next
    
    x = ShellExecute(0, "Open", FileName, "", "", vbNormalNoFocus)

End Function

' ===============================================================
' IsFileOpen
' checks if file is open
' ---------------------------------------------------------------
Function IsFileOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: Error ErrNo
    End Select
End Function

' ===============================================================
' JoinRecordsets
' Joins two recordsets together
' ---------------------------------------------------------------
Function JoinRecordsets(ByVal Rst1 As Recordset, Rst2 As Recordset) As Recordset
    Dim i As Integer
    
    On Error Resume Next
    
    With Rst2
        .MoveFirst
        Do While Not .EOF
            Rst1.AddNew
            
            For i = 0 To .Fields.Count - 1
                Rst1.Fields(i) = Rst2.Fields(i)
            Next
            Rst1.Update
            .MoveNext
        Loop
    End With
    Set JoinRecordsets = Rst1
End Function

' ===============================================================
' TotalLinesInProject
' counts total lines of project
' ---------------------------------------------------------------
Public Sub TotalLinesInProject()
    Dim VBP As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim LineCount As Long
    
    Set VBP = ActiveWorkbook.VBProject
    
    If VBP.Protection = vbext_pp_locked Then
    
    Exit Sub
    End If
    
    For Each VBComp In VBP.VBComponents
        LineCount = LineCount + VBComp.CodeModule.CountOfLines
    Next VBComp
    
    MsgBox "Total lines of code = " & LineCount
End Sub

' ===============================================================
' IsTime
' checks to see whether the passed variable is in time format
' ---------------------------------------------------------------
Public Function IsTime(Expression As Variant) As Boolean
    If IsDate(Expression) Then
        IsTime = (Int(CSng(CDate(Expression))) = 0)
    End If
End Function

' ===============================================================
' TransposeArray
' Transposes input array
' ---------------------------------------------------------------
Public Function TransposeArray(myarray() As Variant) As Variant()
Dim x As Long
Dim y As Long
Dim XLower As Long
Dim XUpper As Long
Dim YLower As Long
Dim YUpper As Long
Dim TempArray As Variant

    XLower = LBound(myarray, 2)
    XUpper = UBound(myarray, 2)
    YLower = LBound(myarray, 1)
    YUpper = UBound(myarray, 1)
    
    ReDim TempArray(XUpper, YUpper)
    For x = XLower To XUpper
        For y = YLower To YUpper
            TempArray(x, y) = myarray(y, x)
        Next y
    Next x
    TransposeArray = TempArray
End Function

' ===============================================================
' GetDocLocalPath
' Gets local path for OneDrive folders
' ---------------------------------------------------------------
Public Function GetDocLocalPath(docPath As String) As String
    Const strcOneDrivePart As String = "https://d.docs.live.net/"
    Dim strRetVal As String, bytSlashPos As Byte
    'return the local path for doc, which is either already a local document or a document on OneDrive
    
    strRetVal = docPath & "\"
    If Left(LCase(docPath), Len(strcOneDrivePart)) = strcOneDrivePart Then 'yep, it's the OneDrive path
        'locate and remove the "remote part"
        bytSlashPos = InStr(Len(strcOneDrivePart) + 1, strRetVal, "/")
        strRetVal = Mid(docPath, bytSlashPos)
        'read the "local part" from the registry and concatenate
        strRetVal = RegKeyRead("HKEY_CURRENT_USER\Environment\OneDrive") & strRetVal
        strRetVal = Replace(strRetVal, "/", "\") 'slashes in the right direction
        strRetVal = Replace(strRetVal, "%20", " ") 'a space is a space once more
    End If
    GetDocLocalPath = strRetVal
    
End Function

' ===============================================================
' RegKeyRead
' Reads registry values
' ---------------------------------------------------------------
Function RegKeyRead(i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function

' ===============================================================
' IsValidEmail
' Ensure the entered email is the correct format
' ---------------------------------------------------------------
Function IsValidEmail(sEmailAddress As String) As Boolean
    'Code from Officetricks
    'Define variables
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    'Use the below regular expressions
    sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Create Regular Expression Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.TEST(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    'Return validation result
    IsValidEmail = bReturn
End Function

' ===============================================================
' CleanString
' leaves only alpha numric chars
' ---------------------------------------------------------------
Function CleanString(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    CleanString = strResult
End Function

' ===============================================================
' CleanSQLText
' ensures text is ready for entering into database with SQL
' ---------------------------------------------------------------
Public Function CleanSQLText(TextInput As Variant, Optional ReturnNULL As Boolean) As Variant
    Dim FindNo As Integer
    Dim i As Integer
    
    If VarType(TextInput) = vbString Then
        If TextInput = "" And ReturnNULL Then
            CleanSQLText = Null
        Else
            Do
                If i = 10000 Then Exit Do
                FindNo = InStr(1, TextInput, "''")
'                Debug.Print FindNo, TextInput
                TextInput = Replace(TextInput, "''", "'")
                i = i + 1
            Loop While FindNo <> 0
            
            CleanSQLText = Replace(TextInput, "'", "''")
        End If
    Else
        CleanSQLText = TextInput
        Exit Function
    End If
    
End Function

Sub SaveShapeAsPicture(FileName As String)
'PURPOSE: Save a selected shape/icon as a PNG file to computer's desktop
'SOURCE: www.thespreadsheetguru.com

Dim cht As ChartObject
Dim ActiveShape As Shape
Dim UserSelection As Variant

'Ensure a Shape is selected
  On Error GoTo NoShapeSelected
    Set UserSelection = ActiveWindow.Selection
    Set ActiveShape = ActiveSheet.Shapes(UserSelection.Name)
  On Error GoTo 0

'Create a temporary chart object (same size as shape)
  Set cht = ActiveSheet.ChartObjects.Add( _
    Left:=ActiveCell.Left, _
    Width:=ActiveShape.Width, _
    Top:=ActiveCell.Top, _
    Height:=ActiveShape.Height)

'Format temporary chart to have a transparent background
  cht.ShapeRange.Fill.Visible = msoFalse
  cht.ShapeRange.Line.Visible = msoFalse
    
'Copy/Paste Shape inside temporary chart
  ActiveShape.Copy
  cht.Activate
  ActiveChart.Paste
  
'Save chart to User's Desktop as PNG File
  cht.Chart.Export GetDocLocalPath(ThisWorkbook.Path) & PICTURES_PATH & FileName

'Delete temporary Chart
  cht.Delete

'Re-Select Shape (appears like nothing happened!)
  ActiveShape.Select

Exit Sub

'ERROR HANDLERS
NoShapeSelected:
  MsgBox "You do not have a single shape selected!"
  Exit Sub

End Sub

' ===============================================================
' CharToAscii
' returns ascii number for a character
' ---------------------------------------------------------------
Public Function CharToAscii(ByVal char As String) As Integer
    Dim ascii As Integer    'the variable that will be returned
    If Len(char) = 0 Then   'if empty string was sent...
        MsgBox "You sent an empty string" '...raise error
        GoTo errFunction
    ElseIf Len(char) > 1 Then   'if char is longer than one character...
        If IsNumeric(char) Then '...if the string is a number, assumed to be ascii code and...
            Select Case CInt(char)
                Case 0 To 31, 127 '...if the number is whitespace in ascii, raise error
                    MsgBox "You have entered a string of numbers." & Chr(10) & _
                    "We assumed you meant to get this string as a ascii code, but this string is code for a whitespace."
                    GoTo errFunction
                Case Is < 0, Is > 127 '...if the number does not exist in ascii, raise error
                    MsgBox "You have entered a string of numbers." & Chr(10) & _
                    "We assumed you meant to get this string as a ascii code, But there is no character for this number."
                    GoTo errFunction
            End Select
            ascii = CInt(char) '...if no error was raise,the number is OK, so return it
        Else                '...if the string is characters...
            MsgBox "you have sent more than one character. We assume you meant the first."
            Select Case Asc(char) '...check the ascii code for whitespace and...
                Case 0 To 31, 127 '...if whitespace, raise error
                    MsgBox "But it was a whitespace."
                    GoTo errFunction
            End Select      'if no error was raised, the number is okay
            ascii = Asc(char) 'asc() returns the ascii for the first character in a string.
        End If
    Else                    ' if only one character was received...
        Select Case Asc(char)
            Case 0 To 31, 127 '...if it is a whitespace, raise error
                MsgBox "You have entered a whitespace."
                GoTo errFunction
        End Select
        ascii = Asc(char)   '...if not return it
    End If
    CharToAscii = ascii
    GoTo finish
errFunction:
    MsgBox "please try again" 'or anything else you want to do in case of error
finish:
End Function
