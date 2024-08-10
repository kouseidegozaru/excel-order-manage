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
Implements SheetAccesserInterface
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnNumber_ As Integer
Private MatchColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private PriceColumnNumber_ As Integer
Private AmountColumnNumber_ As Integer

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputUserCodeRange_ As String
Private OutputUserNameRange_ As String
Private InputTargetDateRange_ As String

Private Sub Class_Initialize()
    WorkSheetName_ = "発注入力"
    
    ProductCodeColumnNumber_ = 1
    MatchColumnNumber_ = 10
    QtyColumnNumber_ = 11
    PriceColumnNumber_ = 13
    AmountColumnNumber_ = 14
    
    InputBumonCodeRange_ = "A2"
    OutputBumonNameRange_ = "B2"
    
    InputUserCodeRange_ = "C2"
    OutputUserNameRange_ = "D2"
    
    InputTargetDateRange_ = "E2"
    
    
    Set super = New SheetAccesser
    super.InitStatus DataHeaderRow:=5, _
                     DataStartRow:=4, _
                     DataStartColumn:=1
                     
    
    InitWorkBook
    InitWorkSheet
End Sub

Private Sub InitWorkBook()
    Set super.InitWorkBook = ThisWorkbook
End Sub

Private Sub InitWorkSheet()
    Set super.InitWorkSheet = super.Workbook.Sheets(WorkSheetName_)
End Sub
Property Get ProductCodeColumnNumber() As Integer
    ProductCodeColumnNumber = ProductCodeColumnNumber_
End Property
Property Get QtyColumnNumber() As Integer
    QtyColumnNumber = QtyColumnNumber_
End Property
Property Get AmountColumnNumber() As Integer
    AmountColumnNumber = AmountColumnNumber_
End Property
Property Get PriceColumnNumber() As Integer
    PriceColumnNumber = PriceColumnNumber_
End Property
Property Get MatchColumnNumber() As Integer
    MatchColumnNumber = MatchColumnNumber_
End Property

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.ProductCodeColumnNumber, _
                                           Me.DataStartRowNumber, _
                                           Me.ProductCodeColumn, _
                                           Me.Me.DataEndRowNumber)
End Property

'発注入力にある商品の合わせ数の範囲
Property Get MatchRange() As Range
    Set MatchRange = super.GetRange(Me.MatchColumnNumber, _
                                    Me.DataStartRowNumber, _
                                    Me.MatchCodeColumnNumber, _
                                    Me.Me.DataEndRowNumber)
End Property

'発注入力にある商品の数量の範囲
Property Get QtyRange() As Range
    Set QtyRange = super.GetRange(Me.QtyColumnNumber, _
                                  Me.DataStartRowNumber, _
                                  Me.QtyCodeColumnNumber, _
                                  Me.Me.DataEndRowNumber)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = super.RangeToOneDimArray(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている合わせ数の取得
Property Get Match() As Collection
    Set Match = super.RangeToOneDimArray(Me.MatchRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = super.RangeToOneDimArray(Me.QtyRange)
End Property

'部門コードの取得
Property Get BumonCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
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

'担当者コードの取得
Property Get UserCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputUserCodeRange_).value
    
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

'対象日付の取得
Property Get TargetDate() As Date
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    '未入力の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    TargetDate = CDate(value)
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
    Set DataHeaderRowIndex = super.DataHeaderRowIndex
End Property

Property Get DataStartRowIndex() As Long
    Set DataStartRowIndex = super.DataStartRowIndex
End Property

Property Get DataStartColumnIndex() As Integer
    Set DataStartColumnIndex = super.DataStartColumnIndex
End Property

Property Get DataEndRowIndex() As Long
    Set DataEndRowIndex = super.DataEndRowIndex
End Property

Property Get AllDataRange() As Range
    Set AllDataRange = super.AllDataRange
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Call super.Cells(rowIndex, columnIndex, value)
End Property

' メソッド
Function DataEndColumnIndex() As Integer
    DataEndColumnIndex = super.DataEndColumnIndex
End Function

Function DataNextRowIndex() As Long
    DataNextRowIndex = super.DataNextRowIndex
End Function

Sub WriteAllData(targetData As Variant)
    Call super.WriteAllData(targetData)
End Sub

Function GetAllData() As Variant
    GetAllData = super.GetAllData
End Function

Function GetAllDara_NoHead() As Variant
    GetAllDara_NoHead = super.GetAllDara_NoHead
End Function
