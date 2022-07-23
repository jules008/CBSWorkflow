Attribute VB_Name = "ModProjectInOut"
'===============================================================
' Module ModProjectInOut
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jul 22
'===============================================================

Option Explicit

' ===============================================================
' ImportModules
' Imports all VBA Modules from dev library
' ---------------------------------------------------------------
Public Sub ImportModules()
    Dim TargetBook As Excel.Workbook
    Dim FSO As Scripting.FileSystemObject
    Dim FileObj As Scripting.File
    Dim TargetBookName As String
    Dim ImportFileName As String
    Dim VBModules As VBIDE.VBComponents
    
    On Error Resume Next
    
    Set FSO = New Scripting.FileSystemObject
    If FSO.GetFolder(ThisWorkbook.Path).Files.Count = 0 Then
       MsgBox "There are no files to import", vbInformation
       Exit Sub
    End If

    Set VBModules = ThisWorkbook.VBProject.VBComponents
    
    For Each FileObj In FSO.GetFolder(ThisWorkbook.Path).Files
        'Debug.Print FileObj.Name
        
        If (FSO.GetExtensionName(FileObj.Name) = "cls") Or _
            (FSO.GetExtensionName(FileObj.Name) = "frm") Or _
            (FSO.GetExtensionName(FileObj.Name) = "bas") And _
            FileObj.Name <> "ModProjectInOut.bas" Then
            VBModules.Import FileObj.Path
        End If
        
    Next FileObj
    'Debug.Print "End of import"
    Set FSO = Nothing
    Set VBModules = Nothing
End Sub
 
' ===============================================================
' CopyShtCodeModule
' Copies sheet modules and this workbook classes
' ---------------------------------------------------------------
Public Sub CopyShtCodeModule()
    Dim SourceMod As VBIDE.VBComponent
    Dim DestMod As VBIDE.VBComponent
    Dim VBModule As VBIDE.VBComponent
    Dim VBCodeMod As VBIDE.CodeModule
    Dim i As Integer

    If ModuleExists("ThisWorkbook1") Then
        Set SourceMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook1")
        Set DestMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook")
    
        If DestMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
        End If
        
        If SourceMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
        End If
    End If
    
    For Each VBModule In ThisWorkbook.VBProject.VBComponents

        With VBModule

            If Left(.Name, 3) = "Sht" And .Type <> vbext_ct_Document Then
                Set SourceMod = VBModule

                For Each DestMod In ThisWorkbook.VBProject.VBComponents
                    If Left(SourceMod.Name, Len(SourceMod.Name) - 1) = DestMod.Name Then

                        If SourceMod.CodeModule.CountOfLines > 0 Then
                            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
    
                            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
                        End If
                    End If
                Next
            End If
        End With
    Next

    For Each VBModule In ThisWorkbook.VBProject.VBComponents
        If Right(VBModule.Name, 1) = "1" And VBModule.Name <> "Sheet1" Then
            ThisWorkbook.VBProject.VBComponents.Remove VBModule
        End If
    Next VBModule
    
    Set SourceMod = Nothing
    Set DestMod = Nothing
    Set VBModule = Nothing
    Set VBCodeMod = Nothing
End Sub

' ===============================================================
' ModuleExists
' checks to see if module exists in project
' ---------------------------------------------------------------
Public Function ModuleExists(ModuleName As String) As Boolean
    Dim CodeModule As VBIDE.VBComponent
 
    For Each CodeModule In ThisWorkbook.VBProject.VBComponents
        If CodeModule.Name = ModuleName Then ModuleExists = True
    Next
End Function



