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

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer
Private MatchColumnIndex_ As Integer
Private QtyColumnIndex_ As Integer
Private PriceColumnIndex_ As Integer
Private AmountColumnIndex_ As Integer

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputUserCodeRange_ As String
Private OutputUserNameRange_ As String
Private InputTargetDateRange_ As String

Private Sub Class_Initialize()
    WorkSheetName_ = "発注入力"
    
    ProductCodeColumnIndex_ = 1
    MatchColumnIndex_ = 10
    QtyColumnIndex_ = 11
    PriceColumnIndex_ = 13
    AmountColumnIndex_ = 14
    
    InputBumonCodeRange_ = "A2"
    OutputBumonNameRange_ = "B2"
    
    InputUserCodeRange_ = "C2"
    OutputUserNameRange_ = "D2"
    
    InputTargetDateRange_ = "E2"
    
    
    Set super = New SheetAccesser
    super.InitStatus DataHeaderRow:=4, _
                     DataStartRow:=5, _
                     DataStartColumn:=1
                     
    
    InitWorkbook
    InitWorkSheet
End Sub

Private Sub InitWorkbook()
    super.InitWorkbook ThisWorkbook
End Sub

Private Sub InitWorkSheet()
    super.InitWorkSheet super.Workbook.Sheets(WorkSheetName_)
End Sub
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property
Property Get qtyColumnIndex() As Integer
    qtyColumnIndex = QtyColumnIndex_
End Property
Property Get AmountColumnIndex() As Integer
    AmountColumnIndex = AmountColumnIndex_
End Property
Property Get priceColumnIndex() As Integer
    priceColumnIndex = PriceColumnIndex_
End Property
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
                                  Me.qtyColumnIndex, _
                                  Me.DataEndRowIndex, _
                                  Me.qtyColumnIndex)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get productsCode() As Collection
    Set productsCode = RangeToOneDimCollection(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている合わせ数の取得
Property Get match() As Collection
    Set match = RangeToOneDimCollection(Me.MatchRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get qty() As Collection
    Set qty = RangeToOneDimCollection(Me.QtyRange)
End Property

'部門コードの取得
Property Get bumonCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    '部門コードが存在するか
    Dim dataStorage As New DataBaseAccesser
    If dataStorage.ExistsBumon(CInt(value)) Then
        bumonCode = CInt(value)
    Else
        bumonCode = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
End Property

'担当者コードの取得
Property Get userCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputUserCodeRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    '担当者コードが存在するか
    Dim dataStorage As New DataBaseAccesser
    If dataStorage.ExistsUser(CInt(value)) Then
        userCode = CInt(value)
    Else
        userCode = 0
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
End Property

'対象日付の取得
Property Get targetDate() As Date
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    targetDate = CDate(value)
End Property

Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property

Property Get UserCodeRange() As Range
    Set UserCodeRange = Me.Worksheet.Range(InputUserCodeRange_)
End Property

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
