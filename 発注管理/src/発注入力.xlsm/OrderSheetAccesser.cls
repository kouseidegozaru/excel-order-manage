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
Private FormatSheetName_ As String
Private FormatSheet_ As Worksheet

Private ProductCodeColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private DataStartRowNumber_ As Long

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputUserCodeRange_ As String
Private OutputUserNameRange_ As String
Private InputTargetDateRange_ As String

Private Sub Class_Initialize()
    WorkBookName_ = "発注入力.xlsm"
    WorkSheetName_ = "発注入力"
    FormatSheetName_ = "商品マスターフォーマット"
    
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
    InitFormatSheet
End Sub

Private Sub InitWorkBook()
    Set WorkBook_ = Workbooks(WorkBookName_)
End Sub

Private Sub InitWorkSheet()
    Set WorkSheet_ = WorkBook_.Sheets(WorkSheetName_)
End Sub

Private Sub InitFormatSheet()
    Set FormatSheet_ = WorkBook_.Sheets(FormatSheetName_)
End Sub

Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
End Property

Property Get FormatSheet() As Worksheet
    Set FormatSheet = FormatSheet_
End Property

Property Get ProductCodeColumnNumber() As Integer
    ProductCodeColumnNumber = ProductCodeColumnNumber_
End Property

Property Get ProductCodeColumn() As String
    ProductCodeColumn = NumberToLetter(Me.ProductCodeColumnNumber)
End Property

Property Get QtyColumnNumber() As Integer
    QtyColumnNumber = QtyColumnNumber_
End Property

Property Get QtyColumn() As String
    QtyColumn = NumberToLetter(Me.QtyColumnNumber)
End Property

Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property

'次に入力する商品情報が空白の行番号
Public Function DataNextRowNumber() As Long
    DataNextRowNumber = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.ProductCodeColumnNumber).End(xlUp).row + 1
End Function

'商品情報が記載されている最後の行番号
Property Get DataEndRowNumber() As Long
    Dim lastRow As Long
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = Me.DataNextRowNumber - 1
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
    Dim value As Variant
    value = Me.Worksheet.range(InputBumonCodeRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    '部門コードが存在するか
    Dim DataStorage As New DataBaseAccesser
    If DataStorage.ExistsBumon(CInt(value)) Then
        BumonCode = CInt(value)
    Else
        BumonCode = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
End Property

Property Get BumonCodeRange() As range
    Set BumonCodeRange = Me.Worksheet.range(InputBumonCodeRange_)
End Property

'担当者コードの取得
Property Get UserCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.range(InputUserCodeRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    '担当者コードが存在するか
    Dim DataStorage As New DataBaseAccesser
    If DataStorage.ExistsUser(CInt(value)) Then
        UserCode = CInt(value)
    Else
        UserCode = 0
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
End Property

Property Get UserCodeRange() As range
    Set UserCodeRange = Me.Worksheet.range(InputUserCodeRange_)
End Property

'対象日付の取得
Property Get TargetDate() As Date
    Dim value As Variant
    value = Me.Worksheet.range(InputTargetDateRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    TargetDate = CDate(value)
End Property

Property Get TargetDateRange() As range
    Set TargetDateRange = Me.Worksheet.range(InputTargetDateRange_)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = GetRangeValue(Me.QtyRange)
End Property

'担当者名の表示
Property Let UserName(name As String)
    Me.Worksheet.range(OutputUserNameRange_).value = name
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.range(OutputBumonNameRange_).value = name
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

Public Sub WriteProductsCode(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.ProductCodeColumnNumber, writeCollection
End Sub

Public Sub WriteQty(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.QtyColumnNumber, writeCollection
End Sub


