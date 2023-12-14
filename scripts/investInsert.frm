VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} investInsert 
   Caption         =   "Insert your new investment"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "investInsert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "investInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()
    autoFillCombo (Me.TextBox1.Text)
End Sub

Private Sub UserForm_Initialize()
    Dim itemTypes As Range
    On Error Resume Next
    Set itemTypes = ThisWorkbook.Names("InvTYPE").RefersToRange
    On Error GoTo 0
    
    If itemTypes Is Nothing Then
        MsgBox "Defined name 'InvTYPE' not found!"
        Me.Hide
        Exit Sub
    End If
    
    Me.ComboBox1.List = itemTypes.value
End Sub

Private Sub CancelBtn_Click()
    Me.TextBox1.Text = ""
    Me.TextBox2.Text = ""
    Me.TextBox3.Text = ""
    Me.TextBox4.Text = ""
    Me.ComboBox1.value = ""
    
    Me.Hide
End Sub

Private Sub ContinueBtn_Click()
    If Me.TextBox1.Text = "" Or Me.TextBox2.Text = "" Or Me.TextBox3.Text = "" Or Me.ComboBox1.value = "" Or Not IsNumeric(Me.TextBox2.Text) Or Not IsNumeric(Me.TextBox3.Text) Then
        MsgBox "Please fill in all required and valid data."
        Exit Sub
    End If
    
    Dim itemName As String: itemName = Me.TextBox1.Text
    Dim itemQty As Double: itemQty = Me.TextBox2.Text
    Dim paidPrice As Double: paidPrice = Me.TextBox3.Text
    Dim itemType As String: itemType = Me.ComboBox1.value
    Dim itemLink As String: itemLink = Me.TextBox4.Text
    Dim priceNow As Double: priceNow = priceScraper(itemLink)
    Dim totVal As Double: totVal = priceNow * itemQty
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CSGO Investments")
    
    Dim table As ListObject
    Set table = ws.ListObjects("InvestTable")
    
    Dim nRow As ListRow
    Set nRow = table.ListRows.Add
    With nRow
        .Range(1) = table.ListRows.Count
        .Range(2) = itemName
        With .Range(3)
            .value = "Link"
            .Hyperlinks.Add Anchor:=.Cells(1), Address:=itemLink
        End With
        .Range(4) = itemType
        .Range(5) = itemQty
        .Range(6) = paidPrice
        .Range(7) = paidPrice / itemQty
        .Range(8) = priceNow
        .Range(9) = totVal
        .Range(10) = (totVal - paidPrice) / paidPrice
    End With
    
    Me.TextBox1.Text = ""
    Me.TextBox2.Text = ""
    Me.TextBox3.Text = ""
    Me.TextBox4.Text = ""
    Me.ComboBox1.value = ""
    
    Me.Hide
End Sub

Private Sub autoFillCombo(itemName As String)
    Dim keywords As Variant
    keywords = Array("Package", "Case", "Capsule", "Sticker", "Factory New", "Minimal Wear", "Battle-Scarred", "Field-Tested", "Well-Worn")
    
    Dim keyword As Variant
    For Each keyword In keywords
        If InStr(1, itemName, keyword, vbTextCompare) > 0 Then
            Select Case keyword
                Case "Case"
                    Me.ComboBox1.value = "Cases"
                Case "Sticker"
                    If InStr(1, itemName, "Capsule", vbTextCompare) > 0 Then
                        Me.ComboBox1.value = "Capsules"
                    Else
                        Me.ComboBox1.value = "Stickers"
                    End If
                Case "Capsule"
                    Me.ComboBox1.value = "Capsules"
                Case "Package"
                    Me.ComboBox1.value = "Packages"
                Case Else
                    Me.ComboBox1.value = "Fillers"
            End Select
        End If
    Next keyword
End Sub

Function priceScraper(link As String) As Double
    Dim httpRequest As Object
    Dim htmlDoc As Object
    Dim items As Object
    Dim item As Object
    Dim className As String
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

Function getCNYPrice() As Double
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
        
        getCNYPrice = CDbl(Replace(result, ".", ","))
        Exit For
    Next item
    
End Function
