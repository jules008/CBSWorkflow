VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmContactForm 
   Caption         =   "Contact"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
   OleObjectBlob   =   "FrmContactForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmContactForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmContactForm
' Admin form for members
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 22
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmContactForm"
Public Event CreateNew()
Public Event Update()
Public Event Delete()

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

' ===============================================================
' BtnDelete_Click
' Marks member as deleted
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to delete the Contact from the database?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
    End If
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtAddress1 = ""
    TxtAddress2 = ""
    TxtContactName = ""
    TxtContactNo = ""
    TxtPhone1 = ""
    TxtPhone2 = ""
    TxtPosition = ""
    CmoContactType = ""
    CmoOrganisation = ""
End Sub

' ===============================================================
' BtnNew_Click
' Creates new Contact
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    RaiseEvent CreateNew
End Sub

' ===============================================================
' BtnUpdate_Click
' update changes and close
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case ValidateForm

        Case Is = enFunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = enValidationError
            
        Case Is = enFormOK
            
            RaiseEvent Update
            Unload Me
    End Select
    
GracefulExit:

Exit Sub

ErrorExit:

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
' PopulateOrgs
' popullates organisations after contact type is selected
' ---------------------------------------------------------------
Private Function PopulateOrgs(ContactType As String) As Boolean
    Dim RstSource As Recordset
    Dim Index As String
    Dim Table As String
    Dim Name As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "populateOrgs()"

    On Error GoTo ErrorHandler

    Select Case ContactType
        Case "Client"
            Index = "ClientNo"
            Table = "TblClient"
            Name = "Name"
        Case "Lender"
            Index = "LenderNo"
            Table = "TblLender"
            Name = "Name"
        Case "SPV"
            Index = "SPVNo"
            Table = "TblSPV"
             Name = "Name"
       Case "Project"
            Index = "ProjectNo"
            Table = "TblProject"
            Name = "ProjectName"
       Case "Lead"
            Index = "ProjectNo"
            Table = "TblProject"
            Name = "ProjectName"
    End Select
        
    If ContactType <> "Lead" Then
    Set RstSource = ModDatabase.SQLQuery("SElECT " & Index & ", " & Name & " FROM " & Table)
    
    With RstSource
        CmoOrganisation.Clear
        Do While Not .EOF
            With CmoOrganisation
                .AddItem
                If Not IsNull(RstSource.Fields(0)) Then .List(i, 0) = RstSource.Fields(0)
                If Not IsNull(RstSource.Fields(1)) Then .List(i, 1) = RstSource.Fields(1)
                i = i + 1
            End With
            .MoveNext
        Loop
    End With
    Else
        CmoOrganisation.Clear
        With CmoOrganisation
            .AddItem
            .List(0, 1) = "None"
        End With
    End If
    
    PopulateOrgs = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    PopulateOrgs = False

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
' TxtContactName_Change
' ---------------------------------------------------------------
Private Sub CmoContactType_Change()
    CmoOrganisation = ""
    With CmoContactType
        If .ListIndex = -1 Then
            CmoOrganisation.Value = "Please select a Contact Type"
        Else
            If Not PopulateOrgs(.Value) Then Err.Raise HANDLED_ERROR
        End If
    End With
End Sub

' ===============================================================
' TxtContactName_Change
' ---------------------------------------------------------------
Private Sub TxtContactName_Change()
    TxtContactName.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtContactNo_Change
' ---------------------------------------------------------------
Private Sub TxtContactNo_Change()
    TxtContactNo.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPhone1_Change
' ---------------------------------------------------------------
Private Sub TxtPhone1_Change()
    TxtPhone1.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPhone2_Change
' ---------------------------------------------------------------
Private Sub TxtPhone2_Change()
    TxtPhone2.BackColor = COL_WHITE
End Sub

' ===============================================================
' TxtPosition_Change
' ---------------------------------------------------------------
Private Sub TxtPosition_Change()
    TxtPosition.BackColor = COL_WHITE
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Me.Tag = "New" Then
        BtnNew.Visible = False
        BtnUpdate.Caption = "Create"
    End If
    ClearForm
    
    With CmoContactType
        .Clear
        .AddItem
        .List(0, 0) = 0
        .List(0, 1) = "Lender"
        .AddItem
        .List(1, 0) = 1
        .List(1, 1) = "Project"
        .AddItem
        .List(2, 0) = 2
        .List(2, 1) = "SPV"
        .AddItem
        .List(3, 0) = 3
        .List(3, 1) = "Client"
        .AddItem
        .List(3, 0) = 4
        .List(3, 1) = "Lead"
    End With
    CmoOrganisation.ForeColor = COL_DRK_GREY
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
    With TxtContactName
        If .Value = "" Then
            .BackColor = COL_AMBER
            ValidateForm = enValidationError
        End If
    End With
                     
    If ValidateForm = enValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = enFormOK
    End If
    
Exit Function

enValidationError:
    
    ValidateForm = enValidationError
Exit Function

ErrorExit:

    ValidateForm = enFunctionalError
Exit Function

ErrorHandler:
    
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        CustomErrorHandler Err.Number
        Resume enValidationError:
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================

