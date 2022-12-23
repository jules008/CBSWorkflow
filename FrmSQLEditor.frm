VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSQLEditor 
   Caption         =   "UserForm1"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "FrmSQLEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSQLEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnClose_Click()
    Hide
End Sub

Private Sub BtnSQLVBA_Click()
    Dim ArySQL() As String
    Dim i As Integer
    
    ArySQL = Split(TxtSQL, vbCrLf)
    
    TxtSQL = ""
    
    For i = LBound(ArySQL) To UBound(ArySQL)
        If i = LBound(ArySQL) Then
            ArySQL(i) = """" & ArySQL(i) & " ""  _"
        ElseIf i = UBound(ArySQL) Then
            ArySQL(i) = " & "" " & ArySQL(i) & """"
        Else
            ArySQL(i) = " & "" " & ArySQL(i) & " ""  _"
        End If
        TxtSQL = TxtSQL & ArySQL(i) & vbNewLine
    Next
        
    
End Sub



