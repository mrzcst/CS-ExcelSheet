VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dataInsert 
   Caption         =   "Insert Data"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "dataInsert.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dataInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim itemTypes As Range
    On Error Resume Next
    Set itemTypes = ThisWorkbook.Names("ItemTYPE").RefersToRange
    On Error GoTo 0
    
    If itemTypes Is Nothing Then
        MsgBox "Defined name 'ItemTYPE' not found!"
        Me.Hide
        Exit Sub
    End If
    
    Me.ComboBox1.List = itemTypes.value
    
    Dim marketNames As Range
    On Error Resume Next
    Set marketNames = ThisWorkbook.Names("MarketNAME").RefersToRange
    On Error GoTo 0
    
    If marketNames Is Nothing Then
        MsgBox "Defined name 'MarketNAME' not found!"
        Me.Hide
        Exit Sub
    End If
    
    Me.ComboBox2.List = marketNames.value
    
    Me.Frame3.Visible = False
    
    Me.ComboBox3.List = WorksheetFunction.Transpose(Array(0, 1, 2, 3, 4, 5, 6, 7, 8))
End Sub

Private Sub ComboBox2_Change()
    If Me.ComboBox2.value = "Skinport" Then
        Me.Frame3.Visible = True
    Else
        Me.Frame3.Visible = False
    End If
End Sub

Private Sub CancelBtn_Click()
    Me.TextBox1.Text = ""
    Me.TextBox2.Text = ""
    Me.ComboBox1.value = ""
    Me.ComboBox2.value = ""
    Me.ComboBox3.value = ""
    
    Me.Hide
End Sub

Private Sub ContinueBtn_Click()
    If Me.TextBox1.Text = "" Or Me.TextBox2.Text = "" Or Me.ComboBox1.value = "" Or Me.ComboBox2.value = "" Or Not IsNumeric(Me.TextBox2.Text) Or (Me.Frame3.Visible And Me.ComboBox3.value = "") Then
        MsgBox "Please fill in all required and valid data."
        Exit Sub
    End If
    
    Dim itemName As String
    Dim paidPrice As Double
    Dim itemType As String
    Dim boughtFrom As String
    Dim tradebleOn As Date
    
    itemName = Me.TextBox1.Text
    paidPrice = CDbl(Replace(Me.TextBox2.Text, ".", ","))
    itemType = Me.ComboBox1.value
    boughtFrom = Me.ComboBox2.value
    
    If boughtFrom = "Buff" Then
        tradebleOn = Date + 8
    ElseIf boughtFrom = "Skinport" Then
        If Me.Frame3.Visible = True Then
            Dim selectedDays As Integer
            selectedDays = Me.ComboBox3.value
            tradebleOn = Date + 8 + selectedDays
        Else
            MsgBox "Please select the tradeble days value."
            Exit Sub
        End If
    Else
        MsgBox "Invalid Bought From value!"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CSGO Trades")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("WaitingList")
    
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    With newRow
        .Range(1) = tbl.ListRows.Count
        .Range(2) = itemName
        .Range(3) = itemType
        .Range(4) = boughtFrom
        .Range(5) = paidPrice
        .Range(6) = tradebleOn
    End With
    
    Me.TextBox1.Text = ""
    Me.TextBox2.Text = ""
    Me.ComboBox1.value = ""
    Me.ComboBox2.value = ""
    Me.ComboBox3.value = ""
    
    Me.Hide
End Sub
