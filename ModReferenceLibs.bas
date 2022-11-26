Attribute VB_Name = "ModReferenceLibs"
'===============================================================
' Module ModReferenceLibs
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 01 Jul 20
'===============================================================

Option Explicit

' ===============================================================
' SetReferenceLibs
' Sets all project reference libraries
' ---------------------------------------------------------------
Public Sub SetReferenceLibs()
    Dim Reference As Object
    
    On Error Resume Next
    
    For Each Reference In ThisWorkbook.VBProject.References
        With Reference
            'Debug.Print .Name
            'Debug.Print .Description
            'Debug.Print .Minor
            'Debug.Print .Major
            'Debug.Print .GUID
            'Debug.Print
        End With
    Next

    ' Visual Basic For Applications
    If Not ReferenceExists("{000204EF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{000204EF-0000-0000-C000-000000000046}", Major:=4, Minor:=1
    End If
    
    ' Microsoft Excel 14.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020813-0000-0000-C000-000000000046}", Major:=1, Minor:=7
    End If
    
    ' Microsoft Forms 2.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0D452EE1-E08F-101A-852E-02608C4D0BB4}", Major:=2, Minor:=0
    End If
    
    ' OLE Automation
    If Not ReferenceExists("{00020430-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020430-0000-0000-C000-000000000046}", Major:=2, Minor:=0
    End If
    
    ' Microsoft Office 14.0 Object Library
    If Not ReferenceExists("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", Major:=2, Minor:=5
    End If
    
    ' Microsoft Office 14.0 Access database engine Object Library
    If Not ReferenceExists("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", Major:=12, Minor:=0
    End If
    
    ' Microsoft Scripting Runtime
    If Not ReferenceExists("{420B2830-E718-11CF-893D-00A0C9054228}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0
    End If
    
    ' Microsoft Visual Basic for Applications Extensibility 5.3
    If Not ReferenceExists("{0002E157-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    End If
    
    ' Microsoft Outlook 14.0 Object Library
    If Not ReferenceExists("{00062FFF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00062FFF-0000-0000-C000-000000000046}", Major:=9, Minor:=4
    End If
    
    ' Adobe Acrobat 10.0 Type Library
    If Not ReferenceExists("{E64169B3-3592-47D2-816E-602C5C13F328}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{E64169B3-3592-47D2-816E-602C5C13F328}", Major:=1, Minor:=1
    End If
    
    ' Microsoft Access 16.0 Object Library
    If Not ReferenceExists("{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", Major:=9, Minor:=0
    End If
    
    ' Microsoft Word 16.0 Object Library
    If Not ReferenceExists("{00020905-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020905-0000-0000-C000-000000000046}", Major:=8, Minor:=7
    End If
End Sub

' ===============================================================
' ReferenceExists
' Checks to see if reference already exists
' ---------------------------------------------------------------
Public Function ReferenceExists(Ref As String) As Boolean
    Dim i As Integer
    
    With ThisWorkbook.VBProject.References
        For i = 1 To .Count
            If .item(i).GUID = Ref Then
                ReferenceExists = True
                Exit Function
            End If
        Next
        ReferenceExists = False
    End With
End Function

