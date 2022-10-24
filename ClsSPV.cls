VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsSPV
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 14 Sep 22
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete - Deletes record
'===============================================================
' Index - SPVNo
'---------------------------------------------------------------
Option Explicit

Private pSPVNo As Integer
Private pName As String
Private pClientNo As Integer
Private pContacts As ClsContacts
Private pProjects As ClsProjects
Private pParent As ClsClient
Private WithEvents pSPVForm As FrmSPVForm
Attribute pSPVForm.VB_VarHelpID = -1

'---------------------------------------------------------------
Public Property Get SPVNo() As Integer
    SPVNo = pSPVNo
End Property

Public Property Let SPVNo(ByVal vNewValue As Integer)
    pSPVNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ClientNo() As Integer
    ClientNo = pClientNo
End Property

Public Property Let ClientNo(ByVal ClientNo As Integer)
    pClientNo = ClientNo
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Contacts() As ClsContacts
    Set Contacts = pContacts
End Property

Public Property Let Contacts(ByRef vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Projects() As ClsProjects
    Set Projects = pProjects
End Property

Public Property Let Projects(ByRef vNewValue As ClsProjects)
    Set pProjects = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsClient
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsClient)
    Set pParent = obj
End Property
'---------------------------------------------------------------
' ===============================================================
' Method DBNew
' displays new SPV form
' ---------------------------------------------------------------
Public Sub DBNew()
    pSPVNo = 0
    pName = ""
    
    With pSPVForm
        .Tag = "New"
        .ClearForm
        .Show
    End With
End Sub

' ===============================================================
' Method DisplayForm
' displays client form
' ---------------------------------------------------------------
Public Sub DisplayForm()
    With pSPVForm
        .Tag = "Update"
        .TxtSPVNo = pSPVNo
        .TxtName = pName
        .Show
    End With
End Sub

' ===============================================================
' Method DBGet(Key as Variant)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(SPVNameOrNo As Variant)
    Dim RstSPV As Recordset

    If IsNumeric(SPVNameOrNo) Then
        Set RstSPV = ModDatabase.SQLQuery("SELECT * FROM TblSPV WHERE SPVNo = " & SPVNameOrNo)
    Else
        Set RstSPV = ModDatabase.SQLQuery("SELECT * FROM TblSPV WHERE Name = '" & SPVNameOrNo & "'")
    End If

    With RstSPV
        If .RecordCount > 0 Then
            If Not IsNull(!SPVNo) Then pSPVNo = !SPVNo
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!ClientNo) Then pClientNo = !ClientNo
            
            pContacts.GetCollection "SPV", pName
        End If
    End With
    Set RstSPV = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstSPV As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstSPV = ModDatabase.SQLQuery("SELECT * FROM TblSPV WHERE SPVNo = " & pSPVNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(SPVNo) FROM TblSPV ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstSPV
        If .RecordCount = 0 Then
            .AddNew
            pSPVNo = LastNo + 1
        Else
            .Edit
        End If

        If CStr(pSPVNo) = "" Then !SPVNo = Null Else !SPVNo = pSPVNo
        If CStr(pName) = "" Then !Name = Null Else !Name = pName
        If CStr(pClientNo) = "" Then !ClientNo = Null Else !ClientNo = pClientNo
        
        .Update
        
        pContacts.SaveCollection
        pProjects.SaveCollection
    End With

    Set RstSPV = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblSPV WHERE SPVNo = " & pSPVNo

End Sub

' ===============================================================
' Event pSPVForm_Update
' Event from client form to update client
' ---------------------------------------------------------------
Private Sub pSPVForm_Update()
    With pSPVForm
        pName = .TxtName
    End With
    DBSave
End Sub

' ===============================================================
' Event pSPVForm_CreateNew
' Event from SPV form to create new SPV
' ---------------------------------------------------------------
Private Sub pSPVForm_CreateNew()
    With pSPVForm
        .Hide
    End With
    DBNew
End Sub

' ===============================================================
' Event pSPVForm_Delete
' Deletes client
' ---------------------------------------------------------------
Private Sub pSPVForm_Delete()
    DBDelete
End Sub

' ===============================================================
Public Sub Initialize(obj As ClsClient)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pContacts = New ClsContacts
    Set pProjects = New ClsProjects
    Set pSPVForm = New FrmSPVForm
    Set pParent = New ClsClient

    pProjects.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pProjects.Terminate

    Set Me.Parent = Nothing

    Set pContacts = Nothing
    Set pProjects = Nothing
    Set pSPVForm = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pContacts = Nothing
    Set pProjects = Nothing
    Set pSPVForm = Nothing
End Sub
' ===============================================================
