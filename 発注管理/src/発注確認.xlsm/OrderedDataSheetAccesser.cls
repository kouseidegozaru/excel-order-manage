VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderedDataSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注済みとしてチェックされた商品コードデータシートへのアクセサ

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

'発注データフォルダパス
Private SaveDirPath_ As String
Private SaveFileName_ As String
'保存データ設定
Private WorkSheetName_ As String

'商品コードのヘッダー名
Private ProductsHeaderName_ As String
'商品コードの列番号
Private ProductCodeColumnIndex_ As Integer

Private Sub Class_Initialize()

    '保存ディレクトリパス
    SaveDirPath_ = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data\ordered"

    'ヘッダー名
    ProductsHeaderName_ = "商品コード"
    
    'シート名
    WorkSheetName_ = "Sheet1"
    
    '商品コード列
    ProductCodeColumnIndex_ = 1
    
    '疑似抽象クラスをインスタンス化
    Set super = New SheetAccesser
    'データのヘッダー行、開始行、開始列を指定
    super.InitStatus DataHeaderRow:=1, _
                     DataStartRow:=2, _
                     DataStartColumn:=1
    
End Sub

'ワークブックを設定
Public Sub InitNewWorkbook()
    ' 新規ワークブックの作成
    super.InitWorkbook Application.Workbooks.Add
End Sub
'ワークブックを設定
Public Sub InitOpenWorkBook()
    ' 既存のワークブックを開く
    super.InitWorkbook Workbooks.Open(SaveFilePath)
End Sub

'ワークシートを設定
Public Sub InitWorkSheet()
    super.InitWorkSheet Me.Workbook.Sheets(WorkSheetName_)
End Sub

'ファイル情報を設定
Public Sub InitStatus(bumonCode As Integer, targetDate As Date)
    SaveFileName_ = "b" & bumonCode & "-" & _
                    "d" & Format(targetDate, "yyyymmdd") & "-" & _
                    ".xlsx"
End Sub

'保存ファイル名
Property Get SaveFileName() As String
    SaveFileName = SaveFileName_
End Property
'保存ファイル名
Property Get SaveDirPath() As String
    SaveDirPath = SaveDirPath_
End Property
'保存ファイルパス
Property Get SaveFilePath() As String
    SaveFilePath = Me.SaveDirPath & "\" & Me.SaveFileName
End Property

'商品コード列番号
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property
'商品コードヘッダー名
Property Get ProductsHeaderName() As String
    ProductsHeaderName = ProductsHeaderName_
End Property
'商品コードの入力
Sub WriteProductsCode(targetData As Collection)
    'ヘッダー入力
    Me.Cells(Me.DataHeaderRowIndex, Me.ProductCodeColumnIndex) = Me.ProductsHeaderName
    '商品コード入力
    WriteOneColumnCollection Me.DataStartRowIndex, Me.ProductCodeColumnIndex, targetData
End Sub

'コレクションを一列に書き込み
Private Sub WriteOneColumnCollection(startRowIndex As Long, startColIndex As Integer, writeData As Collection)
    Dim item As Variant
    Dim ws As Worksheet
    Set ws = Me.Worksheet

    ' writeDataがCollectionか確認
    If TypeName(writeData) = "Collection" Then
        ' Collectionの場合
        rowCount = 0
        colCount = 0
        For Each item In writeData
            ws.Cells(startRowIndex + rowCount, startColIndex + colCount).value = item
            rowCount = rowCount + 1
        Next item
    Else
        ' エラーハンドリング
        Err.Raise vbObjectError + 9999, "writeData", "writeDataはCollectionでなければなりません。"
    End If
End Sub

'ワークブックを保存
Public Sub Save()
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    Me.Workbook.SaveAs fileName:=Me.SaveFilePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
End Sub

'ワークブックを閉じる
Public Sub CloseWorkBook()
    Me.Workbook.Close SaveChanges:=False
End Sub


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




