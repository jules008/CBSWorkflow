VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChkBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Chkbox As MSForms.CheckBox
Attribute Chkbox.VB_VarHelpID = -1

Private Sub ChkBox_Click()
    If Chkbox Then
        Chkbox.Font.Bold = True
    Else
        Chkbox.Font.Bold = False
    End If
End Sub
