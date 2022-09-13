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
' Date - 13 Sep 22
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
Private pContacts As ClsContacts
Private pProjects As ClsProjects
Private pParent As ClsClient

'---------------------------------------------------------------
Public Property Get SPVNo() As Integer
    SPVNo = pSPVNo
End Property

Public Property Let SPVNo(ByVal vNewValue As Integer)
    pSPVNo = vNewValue
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

Public Property Let Contacts(ByVal vNewValue As ClsContacts)
    Set pContacts = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Projects() As ClsProjects
    Set Projects = pProjects
End Property

Public Property Let Projects(ByVal vNewValue As ClsProjects)
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
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(SPVNo As Integer)
    Dim RstSPV As Recordset

    Set RstSPV = ModDatabase.SQLQuery("SELECT * FROM TblSPV WHERE SPVNo = " & SPVNo)
    With RstSPV
        If .RecordCount > 0 Then
            If Not IsNull(!SPVNo) Then pSPVNo = !SPVNo
            If Not IsNull(!Name) Then pName = !Name
        pContacts.GetCollection
        pProjects.GetCollection
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

        pContacts.SaveCollection
        pProjects.SaveCollection
        .Update
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
Public Sub Initialize(obj As ClsClient)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()
    Set pContacts = New ClsContacts
    Set pProjects = New ClsProjects

    pProjects.Initialize Me
End Sub
' ===============================================================
Public Sub Terminate()
    pProjects.Terminate

    Set Me.Parent = Nothing

    Set Contacts = Nothing
    Set Projects = Nothing
End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pContacts = Nothing
    Set pProjects = Nothing
End Sub
' ===============================================================
