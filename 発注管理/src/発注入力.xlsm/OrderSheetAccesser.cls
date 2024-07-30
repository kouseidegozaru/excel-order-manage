VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

Private WorkBookName_ As String
Private WorkBook_ As Workbook

Private WorkSheetName_ As String
Private WorkSheet_ As Worksheet

Private ProductCodeColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private DataStartRowNumber_ As Long

Private InputBumonCDRange_ As String
Private OutputBumonNameRange_  As String
Private InputUserCDRange_ As String
Private OutputUserNameRange_ As String
Private InputDateRange_ As String


Private Sub Class_Initialize()

    WorkBookName_ = "発注入力.xlsm"
    WorkSheetName_ = "発注入力"
    
    ProductCodeColumnNumber_ = 1
    QtyColumnNumber_ = 10
    DataStartRowNumber_ = 5
    
    InputBumonCodeRange_ = "A2"
    OutputBumonNameRange_ = "B2"
    
    InputUserCodeRange_ = "C2"
    OutputUserNameRange_ = "D2"
    
    InputTargetDateRange_ = "E2"
    
    InitWorkBook
    InitWorkSheet
    
End Sub
Private Sub InitWorkBook()
    Dim wb As New Workbook
    Set wb = Workbooks(WorkBookName_)
    Set WorkBook_ = wb
End Sub
Private Sub InitWorkSheet()
    Dim ws As Worksheet
    Set ws = Me.Workbook.Sheets(WorkSheetName_)
    Set WorkSheet_ = wb
End Sub
Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property
Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
End Property
Property Get ProductCodeColumnNumber() As Integer
    ProductCodeColumnNumber = ProductCDColumnNumber_
End Property
Property Get ProductCodeColumn() As String
    ProductCodeColumn = NumberToLetter(Me.ProductCodeColumnNumber)
End Property
Property Get QtyColumnNumber() As Integer
    QtyColumnNumber = ProductCDColumnNumber_
End Property
Property Get QtyColumn() As String
    QtyColumn = NumberToLetter(Me.QtyColumnNumber)
End Property
Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property
'次に入力する商品情報が空白の行番号
Public Function DataNextRowNumer() As Long
    DataNextRowNumer = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.ProductCodeColumnNumber).End(xlUp).row + 1
End Function
'商品情報が記載されている最後の行番号
Property Get DataEndRowNumber() As Long
    Dim lastRow As Long
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = Me.DataNextRowNumer - 1
    If lastRow < Me.DataStartRowNumber Then
        lastRow = Me.DataStartRowNumber
    End If
    DataEndRowNumber = lastRow
End Property

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As range
    
    Dim strRng As String
    strRng = Me.ProductCodeColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.ProductCodeColumn & _
             Me.DataEndRowNumber
    Set ProductsCodeRange = Me.Worksheet.range(strRng)
    
End Property
'発注入力にある商品の数量の範囲
Property Get QtyRange() As range
    
    Dim strRng As String
    strRng = Me.QtyColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.QtyColumn & _
             Me.DataEndRowNumber
    Set QtyRange = Me.Worksheet.range(strRng)
End Property


'部門コードの取得
Property Get BumonCode() As Integer

    '値の取得
    Dim value As Integer
    value = Me.Worksheet.range(InputBumonCodeRange_).value
    
    '部門コードが存在するか
    Dim DataStorage As New DataBaseAccesser
    If DataStorage.ExistsBumon(value) Then
        BumonCode = value
    Else
        BumonCode = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
    
End Property

'担当者コードの取得
Property Get UserCode() As Integer
    
    '値の取得
    Dim value As Integer
    value = Me.Worksheet.range(InputUserCodeRange_).value
    
     '担当者コードが存在するか
    Dim DataStorage As New DataBaseAccesser
    If DataStorage.ExistsUser(value) Then
        UserCode = value
    Else
        UserCode = 0
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
    
End Property

'対象日付の取得
Property Get TargetDate() As Date
    Dim value As Date
    value = Me.Worksheet.range(InputTargetDateRange_).value
    
    TargetDate = value
End Property
'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

'担当者名の表示
Property Let UserName(name As String)
    Me.Worksheet.range(OutputBumonNameRange_).value = name
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.range(OutputUserNameRange_).value = name
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex) = value
End Property

