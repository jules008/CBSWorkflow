VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIDashObjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIDashObjs
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' Julian.turner@OneSheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 21 May 20
'===============================================================

Option Explicit
Private pUIDashObs As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIDashObs.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUIDashObj
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUIDashObs.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UIDashOb As ClsUIDashObj)
    UIDashOb.SetParent ObjPtr(Me.Parent)
    pUIDashObs.Add UIDashOb, Key:=CStr(UIDashOb.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUIDashObs.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIDashObs.Count
End Function

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIFrame
    Dim tmp As ClsUIFrame

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pUIDashObs = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Dim UIDashObj As ClsUIDashObj
    Dim i As Long
    
    For i = pUIDashObs.Count To 1 Step -1
        Set UIDashObj = pUIDashObs(i)
        UIDashObj.SetParent 0
        pUIDashObs.Remove i
        Set UIDashObj = Nothing
    Next i
        
    Set pUIDashObs = Nothing

End Sub

'---------------------------------------------------------------
