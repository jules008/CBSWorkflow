Attribute VB_Name = "ModUIProjects"
'===============================================================
' Module ModUIProjects
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 25 Jun 20
'===============================================================

Option Explicit

Const P_TOP As Integer = 40
Const P_LEFT As Integer = 200
Const P_WIDTH As Integer = 1200
Private ProjExp As Shape
Private projColl As Shape
Private Workflows As Shape
Private CollLWfs As Shape
Private StepDetails As Shape
Private ExpWFs As Shape

Private Const StrMODULE As String = "ModUIProjects"

Public Sub ShowPictures()
    With ShtMain
        Set ProjExp = .Shapes("PRojects_exp")
        Set projColl = .Shapes("Projects_Coll")
        Set Workflows = .Shapes("Workflows")
        Set CollLWfs = .Shapes("CollLWfs")
        Set StepDetails = .Shapes("StepDetails")
        Set ExpWFs = .Shapes("ExpWFs")
    End With

    TogglePicture
End Sub

Public Sub TogglePicture()
    Dim ProjExp As Shape
    Dim projColl As Shape
    Dim Workflows As Shape
    Dim CollLWfs As Shape
    Dim StepDetails As Shape
    Dim ExpWFs As Shape
        
    With ProjExp
        .Top = P_TOP
        .Left = P_LEFT + 1
        .Width = P_WIDTH - 17
    End With
    
    With projColl
        .Top = P_TOP
        .Left = P_LEFT
        .Width = P_WIDTH
    End With
    
    With Workflows
        .Top = 210
        .Left = 230
        .Width = P_WIDTH
    End With
    
    projColl.Visible = msoTriStateToggle
    ProjExp.Visible = msoTriStateToggle
    
    If ProjExp.Visible Then
        Workflows.ZOrder msoBringToFront
        CollLWfs.Visible = msoCTrue
        StepDetails.Visible = msoCTrue
        ExpWFs.Visible = msoFalse
    Else
        Workflows.ZOrder msoSendToBack
        CollLWfs.Visible = msoFalse
        StepDetails.Visible = msoFalse
        ExpWFs.Visible = msoCTrue
    End If
    
    Set ProjExp = Nothing
    Set projColl = Nothing
    Set Workflows = Nothing
    Set CollLWfs = Nothing
    Set StepDetails = Nothing
    Set ExpWFs = Nothing
End Sub

Public Sub WorkflowsClick()
    FrmWorkflow.Show
End Sub

Public Sub HidePictures()
    Workflows.Visible = msoFalse
    CollLWfs.Visible = msoFalse
    StepDetails.Visible = msoFalse
    ExpWFs.Visible = msoFalse
    projColl.Visible = msoFalse
    ProjExp.Visible = msoFalse
   
    Set CollLWfs = Nothing
    Set ProjExp = Nothing
    Set projColl = Nothing
    Set Workflows = Nothing
    Set CollLWfs = Nothing
    Set StepDetails = Nothing
    Set ExpWFs = Nothing

End Sub

