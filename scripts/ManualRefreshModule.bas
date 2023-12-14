Attribute VB_Name = "ManualRefreshModule"
Sub RefreshPricesManual()
    Dim ws As Worksheet
    Dim investTable As ListObject
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("CSGO Investments")
    Set investTable = ws.ListObjects("InvestTable")
    
    For i = 1 To investTable.DataBodyRange.Rows.Count
        Dim link As String
        Dim priceNow As Variant
        Dim qty As Double
        Dim paidPrice As Double
        Dim itemName As String
        
        itemName = investTable.DataBodyRange.Columns(2).Cells(i).value
        link = investTable.DataBodyRange.Columns(3).Cells(i).Hyperlinks(1).Address
        qty = investTable.DataBodyRange.Columns(5).Cells(i).value
        paidPrice = investTable.DataBodyRange.Columns(6).Cells(i).value
        
        Do
            priceNow = InputBox("Enter the price for " & itemName, "Refresh Prices")
            
            If priceNow = "" Then
                Exit Do
            End If
            
            priceNow = Replace(priceNow, ".", ",")
            
            If IsNumeric(priceNow) Then
                investTable.DataBodyRange.Cells(i, 8).value = CDbl(priceNow)
                investTable.DataBodyRange.Cells(i, 9).value = CDbl(priceNow) * qty
                investTable.DataBodyRange.Cells(i, 10).value = ((CDbl(priceNow) * qty) - paidPrice) / paidPrice
                Exit Do
            Else
                MsgBox "Invalid price entered. Please enter a valid numerical value."
            End If
        Loop
    Next i
End Sub

