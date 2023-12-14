Attribute VB_Name = "StateModule"
Sub CheckState()
    Dim isMoveble As Boolean
    isMoveble = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CSGO Trades")
    
    Dim stateTable As ListObject
    Set stateTable = ws.ListObjects("ItemsOnSale")
    
    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Worksheets("Details")
    
    Dim destTable As ListObject
    Set destTable = wsDest.ListObjects("SoldItems")
    
    Dim stateColumn As Range
    Set stateColumn = stateTable.ListColumns("STATE").DataBodyRange
    
    If Not stateTable.DataBodyRange Is Nothing Then
        Dim cell As Range
        For Each cell In stateColumn
            If cell.value = "Sold" Then
                isMoveble = True
                Dim rowNum As Long
                rowNum = cell.Row - stateTable.HeaderRowRange.Row
                MoveSoldItems rowNum, stateTable, destTable
                RenumberItems stateTable
            End If
        Next cell
        If Not isMoveble Then
            MsgBox "There are no sold items!"
        End If
    Else
        MsgBox "No items on sale!"
    End If
End Sub

Private Sub MoveSoldItems(ByVal rowNum As Long, ByVal stateTable As ListObject, ByVal destTable As ListObject)
    Dim itemRow As Range
    Set itemRow = stateTable.DataBodyRange.Rows(rowNum)
    
    Dim firstCell As Range
    Set firstCell = itemRow.Cells(1).Offset(0, 1)
    
    Dim soldPrice As Double
    soldPrice = GetSoldPrice(itemRow.Cells(2).value)
    
    If soldPrice = 0 Then
        MsgBox ("No sell price entered")
        Exit Sub
    End If
    
    Dim destRow As ListRow
    Set destRow = destTable.ListRows.Add
    
    If itemRow.Cells(4).value = "Buff" Then
        soldPrice = soldPrice - (soldPrice * 0.025)
    End If
    
    With destRow
        .Range(1) = destTable.ListRows.Count
        .Range(2) = itemRow.Cells(2).value
        .Range(3) = itemRow.Cells(3).value
        .Range(4) = itemRow.Cells(4).value
        .Range(5) = itemRow.Cells(5).value
        .Range(6) = soldPrice
        .Range(7) = soldPrice - itemRow.Cells(5).value
    End With
    stateTable.ListRows(rowNum).Delete
End Sub

Function GetSoldPrice(ByVal itemName As String) As Double
    On Error Resume Next
    Dim soldPrice As Double
    soldPrice = CDbl(Replace(InputBox("For how much did you sell: " & itemName & "?"), ".", ","))
    
    If Err.Number <> 0 Then
        GetSoldPrice = 0
        Err.Clear
    Else
        GetSoldPrice = soldPrice
    End If
    
    On Error GoTo 0
End Function

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
