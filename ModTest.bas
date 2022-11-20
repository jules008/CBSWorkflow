Attribute VB_Name = "ModTest"

Public Sub TestPicker()
    Dim Picker As ClsFrmPicker
    Dim RstData As Recordset
    
    If DB Is Nothing Then DBConnect
    
    Set Picker = New ClsFrmPicker
    
    Set RstData = ModDatabase.SQLQuery("SELECT Name from TblClient")
    
    With Picker
        .Title = "Select Client"
        .Instructions = "Start typing the name of the Client and select when it appears."
        .ClearForm
        .Data = RstData
        .Show = True
    End With
    MsgBox Picker.SelectedItem
    Stop


    Set Picker = Nothing
    Set RstData = Nothing
End Sub

    
Public Function FormsLoaded()
    Dim Frm As UserForm
    For i = 0 To UserForms.Count - 1
        Debug.Print "Userform " & UserForms(i).Name & " is loaded", UserForms(i).Visible
    Next i
    Debug.Print "Total - " & i
End Function
