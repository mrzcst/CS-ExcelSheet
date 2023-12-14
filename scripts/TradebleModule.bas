Attribute VB_Name = "TradebleModule"
Sub CheckTradeble()
    Dim checkCells As Boolean
    checkCells = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CSGO Trades")
    
    Dim tblSource As ListObject
    Set tblSource = ws.ListObjects("WaitingList")
    
    Dim tblDest As ListObject
    Set tblDest = ws.ListObjects("ItemsOnSale")
    
    Dim tradebleOnColumn As Range
    Set tradebleOnColumn = tblSource.ListColumns("TRADEBLE ON").DataBodyRange
    
    If Not tblSource.DataBodyRange Is Nothing Then
        Dim cell As Range
        For Each cell In tradebleOnColumn
            If cell.value = Date Or cell.value < Date Then
                checkCells = True
                Dim rowNum As Long
                rowNum = cell.Row - tblSource.HeaderRowRange.Row
                MoveTradeble rowNum, tblSource, tblDest
                RenumberItems tblSource
            End If
        Next cell
        If Not checkCells Then
            MsgBox "There are no tradeble items!"
        End If
    Else
        MsgBox "No item in the waiting list!"
    End If
End Sub

Private Sub MoveTradeble(ByVal rowNum As Long, ByVal tblSource As ListObject, ByVal tblDest As ListObject)
    Dim sourceRow As Range
    Set sourceRow = tblSource.DataBodyRange.Rows(rowNum)
    
    Dim firstCell As Range
    Set firstCell = sourceRow.Cells(1).Offset(0, 1)
    
    If firstCell.value <> "" Then
        Set sourceRow = sourceRow.Resize(1, sourceRow.Columns.Count - firstCell.Column + 2)
        
        Dim destRow As ListRow
        Set destRow = tblDest.ListRows.Add
        
        With destRow
            .Range(1) = tblDest.ListRows.Count
            .Range(2) = sourceRow.Cells(2).value
            .Range(3) = sourceRow.Cells(3).value
            .Range(4) = IIf(sourceRow.Cells(4).value = "Buff", "Skinport", "Buff")
            .Range(5) = sourceRow.Cells(5).value
            .Range(6) = "Sellable"
        End With
        
        tblSource.ListRows(rowNum).Delete
    Else
        MsgBox "No data found in the source row."
    End If
End Sub

Sub RenumberItems(ByVal stateTable As ListObject)
    Dim numItems As Long
    numItems = stateTable.ListRows.Count
    
    If numItems > 0 Then
        Dim i As Long
        For i = 1 To numItems
            stateTable.DataBodyRange.Cells(i, 1).value = i
        Next i
    End If
End Sub
