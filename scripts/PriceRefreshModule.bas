Attribute VB_Name = "PriceRefreshModule"
Sub RefreshPrices()
    Dim ws As Worksheet
    Dim investTable As ListObject
    Dim i As Long
    Dim cnyPrice As Double: cnyPrice = scrapeCNYPrice
    
    Set ws = ThisWorkbook.Worksheets("CSGO Investments")
    Set investTable = ws.ListObjects("InvestTable")
    
    For i = 1 To investTable.DataBodyRange.Rows.Count
        Dim link As String
        Dim priceNow As Double
        Dim qty As Double
        Dim paidPrice As Double
        link = investTable.DataBodyRange.Columns(3).Cells(i).Hyperlinks(1).Address
        qty = investTable.DataBodyRange.Columns(5).Cells(i).value
        paidPrice = investTable.DataBodyRange.Columns(6).Cells(i).value
        priceNow = priceScraper(link, cnyPrice)
        
        investTable.DataBodyRange.Cells(i, 8).value = priceNow
        investTable.DataBodyRange.Cells(i, 9).value = priceNow * qty
        investTable.DataBodyRange.Cells(i, 10).value = ((priceNow * qty) - paidPrice) / paidPrice
    Next i
End Sub

Function priceScraper(link As String, getCNYPrice As Double) As Double
    Dim httpRequest As Object
    Dim htmlDoc As Object
    Dim items As Object
    Dim item As Object
    Dim value As String
    Dim price As Double
    
    price = 0
    priceScraper = 0
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", link, False
        .send
    End With
    
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.body.innerHtml = httpRequest.responseText
    
    Set items = htmlDoc.getElementsByClassName("btn btn-default market-button-item")
    
    For Each item In items
        For i = 1 To Len(item.innerText)
            If IsNumeric(Mid(item.innerText, i, 1)) Or Mid(item.innerText, i, 1) = "," Then
                value = value & Mid(item.innerText, i, 1)
            End If
        Next i
        
        If IsNumeric(value) Then
            price = CDbl(Replace(value, ".", ","))
        Else
            MsgBox "Numeric value not found"
        End If
        
        priceScraper = (price * getCNYPrice) * 0.75
        Exit For
    Next item
    
    Set item = Nothing
    Set items = Nothing
    Set htmlDoc = Nothing
    Set httpRequest = Nothing
End Function

Function scrapeCNYPrice() As Double
    Dim httpRequest As Object
    Dim htmlDoc As Object
    Dim items As Object
    Dim item As Object
    Dim link As String
    Dim result As String
    
    Dim startPos As Integer
    Dim endPos As Integer
    
    link = "https://www.currency.me.uk/convert/eur/cny"

    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", link, False
        .send
    End With
    
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.body.innerHtml = httpRequest.responseText
    
    Set items = htmlDoc.getElementsByClassName("mini ccyrate")

    For Each item In items
        
        startPos = InStr(1, item.innerText, "=") + 1
        endPos = InStr(startPos, item.innerText, "CNY") - 1
        result = Trim(Mid(item.innerText, startPos, endPos - startPos + 1))
        
        scrapeCNYPrice = CDbl(Replace(result, ".", ","))
        Exit For
    Next item
    
    Set item = Nothing
    Set items = Nothing
    Set htmlDoc = Nothing
    Set httpRequest = Nothing
End Function
