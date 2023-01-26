VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CmoChangeUser_Change()

    If CurrentUser Is Nothing Then Initialize
        
    If CmoChangeUser.ListIndex <> -1 Then
        CurrentUser.UserLvl = CmoChangeUser
        MsgBox "You now have the following user level - " & CurrentUser.UserLvl, vbOKOnly + vbInformation
    End If
End Sub

Private Sub Worksheet_Activate()
    With CmoChangeUser
        .Clear
        .AddItem "Admin"
        .AddItem "Senior Manager"
        .AddItem "Case Manager"
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
