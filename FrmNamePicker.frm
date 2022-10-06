VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmNamePicker 
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   OleObjectBlob   =   "FrmNamePicker.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmNamePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmNamePicker
' displays form to select names
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 May 20
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmNamePicker"

Public SelMember As ClsMember

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional NewBtn As Boolean) As ClsMember
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    
    If Not IsMissing(NewBtn) Then
        If NewBtn Then BtnNew.Visible = True Else BtnNew.Visible = False
    Else
        BtnNew.Visible = False
    End If
    
    Show
        
    If Not SelMember Is Nothing Then Set ShowForm = SelMember
    Unload Me
    
Exit Function

ErrorExit:
    
    Set SelMember = Nothing

    FormTerminate
    Terminate
    ShowForm = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    If Not SelMember Is Nothing Then Set SelMember = Nothing
    
End Function

' ===============================================================
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    Hide
    FormTerminate
    
End Sub

' ===============================================================
' BtnNew_Click
' Creates new user form
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    Set SelMember = Nothing
    Unload Me
End Sub

' ===============================================================
' BtnSelect_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnSelect_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnSelect_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case ValidateForm

        Case Is = enFunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = enValidationError
            
        Case Is = enFormOK
            FrmNamePicker.Hide
    End Select

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
' TxtSearch_Change
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
    
    TxtSearch.BackColor = COLOUR_8
    LstNames.BackColor = COLOUR_9

    With LstNames
        If .ListIndex <> -1 Then ListResults = .List(.ListIndex)
    
        'if the search box has been changed since being updated by the results box, clear the result box
        If ListResults <> TxtSearch.Value Then .ListIndex = -1
        
        'if the results box has been clicked, add the selected result to the search box
        If .ListIndex = -1 Then
        
            'if no results selected, populate with new results
            If Len(TxtSearch.Value) > 1 Then
                If Not GetSearchItems(TxtSearch.Value) Then Err.Raise HANDLED_ERROR
            Else
                .Clear
            End If
        End If
    End With

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
' ===============================================================
' UserForm_Initialize
' Automatic initialise event that triggers custom Initialise
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    FormInitialise
    
End Sub

' ===============================================================
' UserForm_Terminate
' Automatic Terminate event that triggers custom Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    Set SelMember = New ClsMember
    
    'refresh name list
    If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR

    FormInitialise = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler

    With TxtSearch
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = enValidationError
        End If
    End With
    
    With LstNames
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = enValidationError
        End If
    End With
            
                    
    If ValidateForm = enValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = enFormOK
    End If
    
Exit Function

ValidationError:
    
    ValidateForm = enValidationError

Exit Function

ErrorExit:

    ValidateForm = enFunctionalError
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume ValidationError:
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetSearchItems
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Function GetSearchItems(StrSearch As String) As Boolean
    Const StrPROCEDURE As String = "GetSearchItems()"

    On Error GoTo ErrorHandler

    Dim ListLength As Integer
    Dim RngResult As Range
    Dim StrRange As String
    Dim RngFirstResult As Range
    Dim i As Integer
    
    If IsNumeric(Right(StrSearch, 1)) Then
    
        StrRange = ShtLists.GetSearchRange("Names")
    Else
        StrRange = ShtLists.GetSearchRange("Names")
    
    End If
        
    Set RngResult = ShtLists.Range(StrRange).Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstNames.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
    
        i = 0
        Do
            Set RngResult = ShtLists.Range(StrRange).FindNext(RngResult)
                With LstNames
                    .AddItem
                    If IsNumeric(Right(StrSearch, 1)) Then
                        .List(i, 0) = RngResult.Value
                        .List(i, 1) = RngResult.Offset(0, 1)
                    Else
                        .List(i, 1) = RngResult.Value
                        .List(i, 0) = RngResult.Offset(0, -1)
                    End If
                    i = i + 1
            End With
        Loop While RngResult <> 0 And RngResult.Address <> RngFirstResult.Address
    End If

    GetSearchItems = True
    
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
Exit Function

ErrorExit:

    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
    FormTerminate
    Terminate

    GetSearchItems = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' LstNames_Click
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Sub LstNames_Click()

    On Error Resume Next

    LstNames.BackColor = COLOUR_3
    
    With LstNames
        Me.TxtSearch.Value = .List(.ListIndex, 1)
        .ListIndex = 0
        
        SelMember.DBGet TxtSearch
    End With
    
End Sub



