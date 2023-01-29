VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAccessCntrl 
   Caption         =   "User Access Control Form"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490.001
   OleObjectBlob   =   "FrmAccessCntrl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAccessCntrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmAccessCntrl
' controls access to lenders and Lenders
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jan 23
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAccessCntrl"

Private ActiveUser As ClsCBSUser

' ===============================================================
' ShowForm
' Shows Access control form
' ---------------------------------------------------------------
Public Function ShowForm(CBSUser As ClsCBSUser) As Boolean
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    Set ActiveUser = CBSUser
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    Show


    ShowForm = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    ShowForm = False

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
' PopulateForm
' populates form
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim RstClients As Recordset
    Dim RstLenders As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    Set RstClients = ModDatabase.SQLQuery("Select " _
                                        & "    TblClient.ClientNo, " _
                                        & "    TblClient.Name " _
                                        & "From " _
                                        & "    TblAccessControl Right Join " _
                                        & "    TblClient On TblClient.ClientNo = TblAccessControl.EntityNo " _
                                        & "Where " _
                                        & "    TblAccessControl.Entity = 'Client' And " _
                                        & "    TblAccessControl.UserNo = " _
                                        & ActiveUser.CBSUserNo)

    Set RstLenders = ModDatabase.SQLQuery("Select " _
                                        & "    TblLender.LenderNo, " _
                                        & "    TblLender.Name " _
                                        & "From " _
                                        & "    TblAccessControl Right Join " _
                                        & "    TblLender On TblLender.LenderNo = TblAccessControl.EntityNo " _
                                        & "Where " _
                                        & "    TblAccessControl.Entity = 'Lender' And" _
                                        & "    TblAccessControl.UserNo = " _
                                        & ActiveUser.CBSUserNo)
                                        
    i = 0
    With LstClients
        .Clear
        Do While Not RstClients.EOF
            .AddItem
            .List(i, 0) = RstClients!ClientNo
            .List(i, 1) = RstClients!Name
            i = i + 1
            RstClients.MoveNext
        Loop
    End With
    
    i = 0
    With LstLenders
        .Clear
        Do While Not RstLenders.EOF
            .AddItem
            .List(i, 0) = RstLenders!LenderNo
            .List(i, 1) = RstLenders!Name
            i = i + 1
            RstLenders.MoveNext
        Loop
    End With
    
    TxtUserName = ActiveUser.UserName

    PopulateForm = True
    
    Set RstClients = Nothing
    Set RstLenders = Nothing


Exit Function

ErrorExit:

    Set RstClients = Nothing
    Set RstLenders = Nothing
        
    PopulateForm = False

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
' Form_Intialise
' Form Inititialise
' ---------------------------------------------------------------
Private Function Form_Intialise() As Boolean
    Dim RstSource As Recordset
    
    Const StrPROCEDURE As String = "Form_Intialise()"
    On Error GoTo ErrorHandler
    
    With HdgClients
        .Clear
        .AddItem "Clients"
    End With
    
    With HdgLenders
        .Clear
        .AddItem "Lenders"
    End With
    
    Form_Intialise = True
    
    Set RstSource = Nothing
    
Exit Function

ErrorExit:

    Set RstSource = Nothing
    
    Form_Intialise = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Hide
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As enFormValidation
    
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
           
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
' ClearForm
' Clears form
' ---------------------------------------------------------------
Public Sub ClearForm()
End Sub

' ===============================================================
' UserForm_Initialize
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "UserForm_Initialize()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    If Not Form_Intialise Then Err.Raise HANDLED_ERROR

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
' xBtnAddLender_Click
' Adds Lender to the access list
' ---------------------------------------------------------------
Private Sub xBtnAddLender_Click()
    Dim ErrNo As Integer
    Dim SelLender As Integer
    Dim ActiveLender As ClsLender
    
    Const StrPROCEDURE As String = "xBtnAddLender_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    Dim Picker As ClsFrmPicker
    
    Set Picker = New ClsFrmPicker
    Set ActiveLender = New ClsLender
    
    With Picker
        .Title = "Grant access to Lender"
        .Instructions = "Start typing the name of the Lender and then select from the list. "
        .Data = ModDatabase.SQLQuery("Select  " _
                                    & "     TblLender.Name " _
                                    & " From  " _
                                    & "     TblLender  " _
                                    & " Where  NOT exists   " _
                                    & " (select  " _
                                    & "     TblAccessControl.EntityNo  " _
                                    & " from  " _
                                    & "     TblAccessControl  " _
                                    & " where  " _
                                    & "     TblAccessControl.EntityNo = TblLender.LenderNo and " _
                                    & "     TblAccessControl.Entity = 'Lender' and  " _
                                    & "     TblAccessControl.UserNo = " & ActiveUser.CBSUserNo & ")")
        .ShowNewBtn = False
        .ClearForm
        .Show = True
        
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefulExit
    End If
    
    ActiveLender.DBGet Picker.SelectedItem
    
    If Not ActiveLender Is Nothing Then
    
        DB.Execute "INSERT INTO TblAccessControl (UserNo, Entity, EntityNo) VALUES (" & ActiveUser.CBSUserNo & ", 'Lender', " & ActiveLender.LenderNo & ")"
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    
GracefulExit:
    Set ActiveLender = Nothing
    Set Picker = Nothing

Exit Sub

ErrorExit:

    '***CleanUpCode***
    Set ActiveLender = Nothing
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
' xBtnAddClient_Click
' Adds Client to the access list
' ---------------------------------------------------------------
Private Sub xBtnAddClient_Click()
    Dim ErrNo As Integer
    Dim SelClient As Integer
    Dim ActiveClient As ClsClient
    
    Const StrPROCEDURE As String = "xBtnAddClient_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART

    Dim Picker As ClsFrmPicker
    
    Set Picker = New ClsFrmPicker
    Set ActiveClient = New ClsClient
    
    With Picker
        .Title = "Grant access to Client"
        .Instructions = "Start typing the name of the Client and then select from the list. "
        .Data = ModDatabase.SQLQuery("Select  " _
                                    & "     TblClient.Name " _
                                    & " From  " _
                                    & "     TblClient  " _
                                    & " Where  NOT exists   " _
                                    & " (select  " _
                                    & "     TblAccessControl.EntityNo  " _
                                    & " from  " _
                                    & "     TblAccessControl  " _
                                    & " where  " _
                                    & "     TblAccessControl.EntityNo = TblClient.ClientNo and " _
                                    & "     TblAccessControl.Entity = 'Client' and  " _
                                    & "     TblAccessControl.UserNo = " & ActiveUser.CBSUserNo & ")")
        .ShowNewBtn = False
        .ClearForm
        .Show = True
    
    End With
    
    If Picker.SelectedItem = "" Then
        MsgBox "No selection made, please try again", vbExclamation + vbOKOnly, APP_NAME
        GoTo GracefulExit
    End If
    
    ActiveClient.DBGet Picker.SelectedItem
    
    If Not ActiveClient Is Nothing Then
    
        DB.Execute "INSERT INTO TblAccessControl (UserNo, Entity, EntityNo) VALUES (" & ActiveUser.CBSUserNo & ", 'Client', " & ActiveClient.ClientNo & ")"
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    
GracefulExit:
    Set ActiveClient = Nothing
    Set Picker = Nothing

Exit Sub

ErrorExit:

    '***CleanUpCode***
    Set ActiveClient = Nothing
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
' xBtnRemLender_Click
' Removes Lender from list
' ---------------------------------------------------------------
Private Sub xBtnRemLender_Click()
    Dim ErrNo As Integer
    Dim LenderNo As Integer
    
    Const StrPROCEDURE As String = "xBtnRemLender_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    With LstLenders
        If .ListIndex = -1 Then GoTo GracefulExit
        
        LenderNo = .List(.ListIndex, 0)
        DB.Execute "DELETE * FROM TblAccessControl WHERE UserNo = " & ActiveUser.CBSUserNo _
                    & " AND Entity = 'Lender' " _
                    & " AND EntityNo = " & LenderNo
        
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        
    End With
    
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
' xBtnRemClient_Click
' Removes Client from list
' ---------------------------------------------------------------
Private Sub xBtnRemClient_Click()
    Dim ErrNo As Integer
    Dim ClientNo As Integer
    
    Const StrPROCEDURE As String = "xBtnRemClient_Click()"

    On Error GoTo ErrorHandler

Restart:

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    With LstClients
        If .ListIndex = -1 Then GoTo GracefulExit
        
        ClientNo = .List(.ListIndex, 0)
        DB.Execute "DELETE * FROM TblAccessControl WHERE UserNo = " & ActiveUser.CBSUserNo _
                    & " AND Entity = 'Client' " _
                    & " AND EntityNo = " & ClientNo
        
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        
    End With
    
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





