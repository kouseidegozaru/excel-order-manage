VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer
Private QtyColumnIndex_ As Integer
Private PriceColumnIndex_ As Integer
Private AmountColumnIndex_ As Integer

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputTargetDateRange_ As String
Private InputUserCodeRange_ As String
Private OutputUserNameRange_ As String

Private FilePropertyStartColumnIndex_ As Integer
Private FilePropertyUserNameRowIndex_ As Integer
Private FilePropertyCompleteRowIndex_ As Integer
Private FilePropertyUpdateTimeRowIndex_ As Integer

Private CompleteString_ As String

Private Sub Class_Initialize()
    WorkSheetName_ = "発注確認"
    
    CompleteString_ = "◎"
    
    ProductCodeColumnIndex_ = 1
    QtyColumnIndex_ = 11
    PriceColumnIndex_ = 13
    AmountColumnIndex_ = 14
    
    FilePropertyStartColumnIndex_ = 2
    FilePropertyUserNameRowIndex_ = 4
    FilePropertyCompleteRowIndex_ = 5
    FilePropertyUpdateTimeRowIndex_ = 6
    
    InputBumonCodeRange_ = "A2"
    OutputBumonNameRange_ = "B2"
    
    InputTargetDateRange_ = "C2"
    
    Set super = New SheetAccesser
    super.InitStatus DataHeaderRow:=8, _
                     DataStartRow:=9, _
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

Property Get QtyColumnIndex() As Integer
    QtyColumnIndex = QtyColumnIndex_
End Property

Property Get AmountColumnIndex() As Integer
    AmountColumnIndex = AmountColumnIndex_
End Property

Property Get PriceColumnIndex() As Integer
    PriceColumnIndex = PriceColumnIndex_
End Property

Property Get FilePropertyStartColumnIndex() As Integer
    FilePropertyStartColumnIndex = FilePropertyStartColumnIndex_
End Property

'次に入力するファイル情報が空白の列番号
Public Function FilePropertyNextColumnIndex() As Integer
    FilePropertyNextColumnIndex = Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, Me.Worksheet.Columns.Count).End(xlToLeft).Column + 1
End Function

'ファイル情報が記載されている最後の列番号
Property Get FilePropertyEndColumnIndex() As Long
    Dim lastColumn As Long
    'ファイル情報にデータがない場合列範囲をずらす(無いと範囲にヘッダー列も含まれてしまう)
    lastColumn = Me.FilePropertyNextColumnIndex - 1
    If lastColumn < Me.FilePropertyStartColumnIndex Then
        lastColumn = Me.FilePropertyStartColumnIndex
    End If
    FilePropertyEndColumnIndex = lastColumn
End Property

'ファイル情報の範囲
Property Get FilePropertyRange() As Range
    Set FilePropertyRange = super.GetRange(FilePropertyUserNameRowIndex_, _
                                           Me.FilePropertyStartColumnIndex, _
                                           FilePropertyUpdateTimeRowIndex_, _
                                           Me.FilePropertyEndColumnIndex)
End Property

Public Sub ClearFileProperty()
    Me.FilePropertyRange.Delete Shift:=xlToLeft
End Sub

Public Sub AddFileProperty(UserName As String, IsCompleted As Boolean, updatedTime As Date)
    Dim lastColumn As Integer
    lastColumn = FilePropertyNextColumnIndex
    
    Dim writeCompleteString As String
    writeCompleteString = ""
    If IsCompleted Then
        writeCompleteString = CompleteString_
        
        '更新日時書き込み
        Me.Worksheet.Cells(FilePropertyUpdateTimeRowIndex_, lastColumn) = updatedTime
    End If
    
    Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, lastColumn) = UserName
    Me.Worksheet.Cells(FilePropertyCompleteRowIndex_, lastColumn) = writeCompleteString
End Sub

Public Sub ClearData()
    Me.ProductsCodeRange.EntireRow.Delete Shift:=xlToup
End Sub

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.ProductCodeColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.ProductCodeColumnIndex)
End Property

'発注入力にある商品の数量の範囲
Property Get QtyRange() As Range
    Set QtyRange = super.GetRange(Me.DataStartRowIndex, _
                                  Me.QtyColumnIndex, _
                                  Me.DataEndRowIndex, _
                                  Me.QtyColumnIndex)
End Property
'部門コードの取得
Property Get BumonCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    '空の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    '部門コードが存在するか
    Dim DataStorage As New DataBaseAccesser
    If DataStorage.ExistsBumon(CInt(value)) Then
        BumonCode = value
    Else
        BumonCode = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
End Property

'対象日付の取得
Property Get targetDate() As Date
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    '空の場合終了
    If IsEmpty(value) Then
        End
    End If
    
    targetDate = CDate(value)
End Property

Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property

Property Get TargetDateRange() As Range
    Set TargetDateRange = Me.Worksheet.Range(InputTargetDateRange_)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = RangeToOneDimCollection(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = RangeToOneDimCollection(Me.QtyRange)
End Property

'担当者名の表示
Property Let UserName(name As String)
    Me.Worksheet.Range(OutputUserNameRange_).value = name
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.Range(OutputBumonNameRange_).value = name
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

Public Function AllGroupData() As ADODB.Recordset
    Dim dataAccessor As New SqlExecutor
    Dim rs As ADODB.Recordset
    Dim FilePath As String
    Dim sheetName As String
    Dim rangeName As String
    Dim query As String
    
    ' Excel情報
    FilePath = Me.Workbook.FullName
    sheetName = Me.Worksheet.name
    rangeName = Replace(Me.DataRange.Address(False, False), "$", "")
    
    
    ' 実行するクエリ
    query = "SELECT 品CD, 入荷時間帯, 仕入先名, 発注先名, メーカー名, 品名, 規格, 計区, 入数, 合わせ数, SUM(数量) AS 合計数量, 荷姿, 仕入単価, 仕入金額, LT, 商品備考, メーカー読み " & _
           "FROM [" & sheetName & "$" & rangeName & "] " & _
           "GROUP BY 品CD, 入荷時間帯, 仕入先名, 発注先名, メーカー名, 品名, 規格, 計区, 入数, 合わせ数, 荷姿, 仕入単価, 仕入金額, LT, 商品備考, メーカー読み"
    
    ' Excelファイルにクエリを実行
    Set rs = dataAccessor.ExecuteExcelQuery(FilePath, query)

    Set ProductsGroupData = rs
End Function


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

