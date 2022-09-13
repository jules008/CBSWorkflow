VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsProject
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
' Index - ProjectNo
'---------------------------------------------------------------
Option Explicit

Private pProjectNo As Integer
Private pLoanTerm As String
Private pCBSComPC As Single
Private pExitFee As Boolean
Private pParent As ClsSPV

'---------------------------------------------------------------
Public Property Get ProjectNo() As Integer
    ProjectNo = pProjectNo
End Property

Public Property Let ProjectNo(ByVal vNewValue As Integer)
    pProjectNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LoanTerm() As String
    LoanTerm = pLoanTerm
End Property

Public Property Let LoanTerm(ByVal vNewValue As String)
    pLoanTerm = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CBSComPC() As Single
    CBSComPC = pCBSComPC
End Property

Public Property Let CBSComPC(ByVal vNewValue As Single)
    pCBSComPC = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ExitFee() As Boolean
    ExitFee = pExitFee
End Property

Public Property Let ExitFee(ByVal vNewValue As Boolean)
    pExitFee = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsSPV
    Set Parent = pParent
End Property

Public Property Set Parent(obj As ClsSPV)
    Set pParent = obj
End Property
'---------------------------------------------------------------
' ===============================================================
' Method DBGet(Key as Integer)
' Gets class from Database
' ---------------------------------------------------------------
Public Sub DBGet(ProjectNo As Integer)
    Dim RstProject As Recordset

    Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & ProjectNo)
    With RstProject
        If .RecordCount > 0 Then
            If Not IsNull(!ProjectNo) Then pProjectNo = !ProjectNo
            If Not IsNull(!LoanTerm) Then pLoanTerm = !LoanTerm
            If Not IsNull(!CBSComPC) Then pCBSComPC = !CBSComPC
            If Not IsNull(!ExitFee) Then pExitFee = !ExitFee
        End If
    End With
    Set RstProject = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
' ---------------------------------------------------------------
Public Sub DBSave()
    Dim RstProject As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstProject = ModDatabase.SQLQuery("SELECT * FROM TblProject WHERE ProjectNo = " & pProjectNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(ProjectNo) FROM TblProject ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstProject
        If .RecordCount = 0 Then
            .AddNew
            pProjectNo = LastNo + 1
        Else
            .Edit
        End If

        If CStr(pProjectNo) = "" Then !ProjectNo = Null Else !ProjectNo = pProjectNo
        If CStr(pLoanTerm) = "" Then !LoanTerm = Null Else !LoanTerm = pLoanTerm
        If CStr(pCBSComPC) = "" Then !CBSComPC = Null Else !CBSComPC = pCBSComPC
        If CStr(pExitFee) = "" Then !ExitFee = Null Else !ExitFee = pExitFee

        .Update
    End With

    Set RstProject = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete
' Deletes Record
' ---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)

    DB.Execute "DELETE * FROM TblProject WHERE ProjectNo = " & pProjectNo

End Sub
' ===============================================================
Public Sub Initialize(obj As ClsSPV)
    Set Me.Parent = obj
End Sub
'---------------------------------------------------------------

Private Sub Class_Initialize()

End Sub
' ===============================================================
Public Sub Terminate()

    Set Me.Parent = Nothing

End Sub
'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub
' ===============================================================
