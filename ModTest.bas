Attribute VB_Name = "ModTest"

Public Sub TestSteps()
    Dim Workflow As ClsWorkflow
    Dim RstWorkflow As Recordset
    
    Set Workflow = New ClsWorkflow
    Set RstWorkflow = ModDatabase.SQLQuery("SELECT * FROM TblWorkflow WHERE WorkflowNo > 23")
        
    With RstWorkflow
        Do While Not .EOF
            Set Workflow = New ClsWorkflow
            Workflow.DBGet !WorkflowNo
            Debug.Print "Workflow"; !WorkflowNo, "Progress"; Workflow.Progress
            Workflow.DBSave
            Set Workflow = Nothing
            DoEvents
            .MoveNext
        Loop
    End With
    
    Set RstWorkflow = Nothing
    Set Workflow = Nothing
End Sub
