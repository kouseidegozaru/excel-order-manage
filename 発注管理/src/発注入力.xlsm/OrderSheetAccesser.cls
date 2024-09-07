VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力シートのアクセサ

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer '品コード
Private PiecesColumnIndex_ As Integer '入数
Private MatchColumnIndex_ As Integer '合わせ数
Private QtyColumnIndex_ As Integer '数量
Private PriceColumnIndex_ As Integer '単価
Private AmountColumnIndex_ As Integer '仕入金額

Private InputBumonCodeRange_ As String '部門コードの入力セル
Private OutputBumonNameRange_ As String '部門名の出力セル
Private InputUserCodeRange_ As String '担当者コードの入力セル
Private OutputUserNameRange_ As String '担当者名の出力セル
Private InputTargetDateRange_ As String '発注日の入力セル

Private Sub Class_Initialize()
    'シート名
    WorkSheetName_ = "発注入力"
    
    '商品コードの列
    ProductCodeColumnIndex_ = 1
    '入数の列
    PiecesColumnIndex_ = 9
    '合わせ数の列
    MatchColumnIndex_ = 10
    '数量の列
    QtyColumnIndex_ = 11
    '単価の列
    PriceColumnIndex_ = 13
    '合計金額の列
    AmountColumnIndex_ = 14
    
    '部門コードの入力セル
    InputBumonCodeRange_ = "A2"
    '部門名の出力セル
    OutputBumonNameRange_ = "B2"
    '担当者コードの入力セル
    InputUserCodeRange_ = "C2"
    '担当者名の出力セル
    OutputUserNameRange_ = "D2"
    '発注日の入力セル
    InputTargetDateRange_ = "E2"
    
    
    '疑似抽象クラスをインスタンス化
    Set super = New SheetAccesser
    'データのヘッダー行、開始行、開始列を指定
    super.InitStatus DataHeaderRow:=4, _
                     DataStartRow:=5, _
                     DataStartColumn:=1
    InitWorkbook
    InitWorkSheet
End Sub

'ワークブックを設定
Private Sub InitWorkbook()
    super.InitWorkbook ThisWorkbook
End Sub

'ワークシートを設定
Private Sub InitWorkSheet()
    super.InitWorkSheet super.Workbook.Sheets(WorkSheetName_)
End Sub

'商品コードの列番号
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property
'入数の列番号
Property Get PiecesColumnIndex() As Integer
    PiecesColumnIndex = PiecesColumnIndex_
End Property
'数量の列番号
Property Get QtyColumnIndex() As Integer
    QtyColumnIndex = QtyColumnIndex_
End Property
'合計金額の列番号
Property Get AmountColumnIndex() As Integer
    AmountColumnIndex = AmountColumnIndex_
End Property
'単価の列番号
Property Get PriceColumnIndex() As Integer
    PriceColumnIndex = PriceColumnIndex_
End Property
'合わせ数の列番号
Property Get MatchColumnIndex() As Integer
    MatchColumnIndex = MatchColumnIndex_
End Property

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.ProductCodeColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.ProductCodeColumnIndex)
End Property

'発注入力にある商品の合わせ数の範囲
Property Get MatchRange() As Range
    Set MatchRange = super.GetRange(Me.DataStartRowIndex, _
                                    Me.MatchColumnIndex, _
                                    Me.DataEndRowIndex, _
                                    Me.MatchColumnIndex)
End Property

'発注入力にある商品の数量の範囲
Property Get QtyRange() As Range
    Set QtyRange = super.GetRange(Me.DataStartRowIndex, _
                                  Me.QtyColumnIndex, _
                                  Me.DataEndRowIndex, _
                                  Me.QtyColumnIndex)
End Property

'発注入力に入力されている商品コードの取得
Property Get productsCode() As Collection
    Set productsCode = RangeToOneDimCollection(Me.ProductsCodeRange)
End Property

'発注入力に入力されている合わせ数の取得
Property Get match() As Collection
    Set match = RangeToOneDimCollection(Me.MatchRange)
End Property

'発注入力に入力されている数量の取得
Property Get qty() As Collection
    Set qty = RangeToOneDimCollection(Me.QtyRange)
End Property

'部門コードの取得
Property Get bumonCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    'バリデーションチェックを行い失敗したら終了
    CheckExistsBumon value
    
    bumonCode = CInt(value)

End Property

'担当者コードの取得
Property Get userCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputUserCodeRange_).value

    'バリデーションチェックを行い失敗したら終了
    CheckExistsUser value
    
    userCode = CInt(value)

End Property

'対象日付の取得
Property Get targetDate() As Date
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    'バリデーションチェックを行い失敗したら終了
    CheckDateFormat value
    
    targetDate = CDate(value)
End Property

'部門コードの範囲
Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property
'担当者コードの範囲
Property Get UserCodeRange() As Range
    Set UserCodeRange = Me.Worksheet.Range(InputUserCodeRange_)
End Property
'発注日の範囲
Property Get TargetDateRange() As Range
    Set TargetDateRange = Me.Worksheet.Range(InputTargetDateRange_)
End Property

'担当者名の表示
Property Let UserName(name As String)
    Me.Worksheet.Range(OutputUserNameRange_).value = name
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.Range(OutputBumonNameRange_).value = name
End Property


'''継承部分'''

'プロパティ
Property Get Workbook() As Workbook
    Set Workbook = super.Workbook
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = super.Worksheet
End Property

Property Get DataHeaderRowIndex() As Long
    DataHeaderRowIndex = super.DataHeaderRowIndex
End Property

Property Get DataStartRowIndex() As Long
    DataStartRowIndex = super.DataStartRowIndex
End Property

Property Get DataStartColumnIndex() As Integer
    DataStartColumnIndex = super.DataStartColumnIndex
End Property

Property Get DataEndRowIndex() As Long
    DataEndRowIndex = super.DataEndRowIndex
End Property

Property Get AllDataRange() As Range
    Set AllDataRange = super.AllDataRange
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    super.Cells(rowIndex, columnIndex) = value
End Property

' メソッド
Function DataEndColumnIndex() As Integer
    DataEndColumnIndex = super.DataEndColumnIndex
End Function

Function DataNextRowIndex() As Long
    DataNextRowIndex = super.DataNextRowIndex
End Function

Sub WriteAllData(targetData As Collection)
    Call super.WriteAllData(targetData)
End Sub

Function GetAllData() As Collection
    Set GetAllData = super.GetAllData
End Function

Function GetAllData_NoHead() As Collection
    Set GetAllData_NoHead = super.GetAllData_NoHead
End Function
