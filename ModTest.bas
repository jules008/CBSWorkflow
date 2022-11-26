Attribute VB_Name = "ModTest"
' ===============================================================
' Method BuildTable
' Builds matrix of cells once dimensions are received.
'---------------------------------------------------------------
Public Sub BuildTable(Optional SplitRow As Integer, Optional SplitSize As Integer)
    Dim x, y As Integer
    Dim Cell As ClsUICell
    Dim ColOffset As Integer
    Dim RowOffset As Integer
    Static SplitFlag As Boolean

        For y = 1 To pNoRows
            If SplitRow <> 0 Then
    
                'last row?
                If y < pNoRows Then
                    If SplitFlag Then
    
                    End If
                    
                    'do we split rows next
                    If y = SplitRow Then SplitFlag = True Else SplitFlag = False
                End If
            End If
                    
            For x = 0 To pNoCols - 1
                Set Cell = New ClsUICell
                Cell.Name = "Cell - " & x & "-" & y
                pCells.AddCell Cell, x, y
                With Cell
                    .Left = .Parent.Left + ColOffset
                    .Top = .Parent.Top + RowOffset
                    .Width = pColWidths(x)
                    
                    If x = 0 And ExpandIcon <> "" Then
                        .Badge = pShpExpand.Duplicate
                        With .Badge
                            If y = SplitRow Then
                                .Rotation = 90
                                .Top = Cell.Top + 3
                                .Left = Cell.Left + 30
                                .Height = 13
                                .Width = 11
                                .OnAction = pOnAction(x, y - 1)
                            Else
                                .Top = Cell.Top + 3
                                .Left = Cell.Left + 20
                                .Height = 13
                                .Width = 11
                                .OnAction = pOnAction(x, y - 1)
                            End If
                        End With
                    End If

                    If y = 0 Then
                        .Text = pHeadingText(x)
                    Else
                        'cater for blank columns with no data
                        If x < pRstText.Fields.Count Then
                        If Not IsNull(pRstText.Fields(x)) Then .Text = pRstText.Fields(x)
                        .OnAction = pOnAction(x, y - 1)
                        End If
                    End If
                    
                    If pStyles(x, y - 1) = "" Then
                        If pStyles(1, y - 1) = "" Then
                            .Style = pStylesColl.Find(pStyles(1, 2))
                        Else
                            .Style = pStylesColl.Find(pStyles(1, y))
                        End If
                    Else
                        .Style = pStylesColl.Find(pStyles(x, y - 1))
                    End If
            
                    .Height = pRowHeight
                End With
                ColOffset = ColOffset + Cell.Width + HPad
            Next
            pRstText.MoveNext
            ColOffset = 0
            RowOffset = RowOffset + Cell.Height + VPad
    
    Next
    End If
End Sub
