Attribute VB_Name = "ModProfiler"
'===============================================================
' Module ModProfiler
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Jul 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModProfiler"

' ===============================================================
' ObjectCheck
' Checks objects are closed down correctly
' ---------------------------------------------------------------
Public Sub ObjectCheck()
    Dim VBModule As VBIDE.VBComponent
    Dim i, X As Integer
    Dim NoObjs As Single
    Dim NoObjsClsd As Single
    Dim LineCode As String
    Dim TmpAry() As String
    Dim Objs() As String
    Dim ObjCnt As Integer
    
    For Each VBModule In ThisWorkbook.VBProject.VBComponents
        ReDim Objs(1)
        For i = 1 To VBModule.CodeModule.CountOfLines
            With VBModule.CodeModule
                LineCode = .Lines(i, 1)
                LineCode = Replace(LineCode, vbTab, "")
                LineCode = Trim(LineCode)
                
                Debug.Print LineCode
                If InStr(1, LineCode, "Set ", vbTextCompare) Then
                    TmpAry = Split(LineCode, " ")
                    
                    If InStr(1, LineCode, "Nothing", vbTextCompare) Then
                        For X = LBound(Objs) To UBound(Objs)
                            If TmpAry(1) = Objs(X) Then
                                NoObjsClsd = NoObjsClsd + 1
                                Objs(X) = ""
                                Exit For
                            End If
                        Next
                        
                    Else
                        Objs(ObjCnt) = TmpAry(1)
                        ObjCnt = ObjCnt + 1
                        ReDim Preserve Objs(ObjCnt)
                        NoObjs = NoObjs + 1
                    End If
                End If
            End With
            Debug.Print "Module - " & VBModule.Name
            Debug.Print "No of Objs opened- " & NoObjs
            Debug.Print "No of Objs closed - " & NoObjsClsd
            Debug.Print "Objects left open" & vbCr
            Debug.Print "=================" & vbCr
            
            For X = LBound(Objs) To UBound(Objs)
                If Objs(X) <> "" Then Debug.Print Objs(X) & vbCr
            Next
        Next
    Next
End Sub
