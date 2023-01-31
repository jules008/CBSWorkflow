VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProjConts 
   Caption         =   "Project Contacts"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
   OleObjectBlob   =   "FrmProjConts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProjConts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmProjConts
' Admin form for members
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 31 Jan 23
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmProjConts"

Private SelContact As Integer

Public Event Display(ContactNo As Integer)
Public Event Add()
Public Event Edit(ContactNo As Integer)
Public Event Delete(ContactNo As Integer)

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

'===============================================================
' BtnUpdate_Click
'---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    With LstContacts
        If BtnUpdate.Caption = "Update" Then
            If .ListIndex <> -1 Then
                RaiseEvent Edit(.List(.ListIndex, 0))
            End If
        End If
        
        If BtnUpdate.Caption = "Save" Then
            RaiseEvent Add
            BtnUpdate.Caption = "Update"
        End If
    End With
End Sub

' ===============================================================
' LstContacts_Click
' displays contact details when clicked
' ---------------------------------------------------------------
Private Sub LstContacts_Click()
    With LstContacts
        RaiseEvent Display(.List(.ListIndex, 0))
    End With
End Sub

' ===============================================================
' xBtnNew_Click
' Adds contact
' ---------------------------------------------------------------
Private Sub xBtnNew_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "xBtnNew_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
        ClearForm
        BtnUpdate.Caption = "Save"
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
' xBtnDelete_Click
' Deletes contact
' ---------------------------------------------------------------
Private Sub xBtnDelete_Click()
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    Response = MsgBox("Are you sure you want to delete the Contact from the project?", vbYesNo + vbExclamation, APP_NAME)
    
    If Response = 6 Then
    
        With LstContacts
            If .ListIndex = -1 Then GoTo GracefulExit
            
            SelContact = .List(.ListIndex, 0)
            RaiseEvent Delete(SelContact)
            ClearForm
        End With
    End If
GracefulExit:

Exit Sub


ErrorHandler:
    Dim ErrNo As Integer
    If Err.Number >= 2000 And Err.Number <= 2500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
    End If
End Sub

' ===============================================================
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
    TxtContactName = ""
    TxtEmailAddress = ""
    TxtOrganisation = ""
    TxtPhone1 = ""
    TxtPhone2 = ""
    TxtPosition = ""
    xTxtNotes = ""
End Sub

' ===============================================================
' UserForm_Initialize
' Initialises form controls
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ClearForm
    
    With HdgContacts
        .Clear
        .AddItem
        .List(0, 0) = "Name"
        .List(0, 1) = "Organisation"
    End With
    
End Sub

