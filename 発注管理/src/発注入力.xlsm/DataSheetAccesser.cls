VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

'発注データフォルダパス
Private OrderDataDirPath As String
'保存データ設定

Private WorkBook_ As Workbook

Private WorkSheetName_ As String
Private WorkSheet_ As Worksheet

Private ProductCodeColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private DataStartColumnNumber_ As Integer
Private DataStartRowNumber_ As Long

Private Sub Class_Initialize()

    OrderDataDirPath = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data"

    WorkSheetName_ = "Sheet1"
    
    ProductCodeColumnNumber_ = 1
    QtyColumnNumber_ = 2
    DataStartColumnNumber_ = 1
    DataStartRowNumber_ = 1
    
End Sub

'保存ファイル名
Property Get SaveFileName() As String
    Dim order As New OrderSheetAccesser
    
    SaveFileName = "b" & order.BumonCode & "-" & _
                  "u" & order.UserCode & "-" & _
                  "d" & Format(order.TargetDate, "yyyymmdd") & "-" & _
                  ".xlsx"
End Property

Public Sub InitWorkSheet()
    Set WorkSheet_ = Me.Workbook.Sheets(WorkSheetName_)
End Sub

Public Sub NewWorkbook()
    ' 新規ワークブックの作成
    Set Me.Workbook = Application.Workbooks.add
End Sub

Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property

Property Set Workbook(wb As Workbook)
    Set WorkBook_ = wb
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
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

Property Get DataStartColumnNumber() As Integer
    DataStartColumnNumber = DataStartColumnNumber_
End Property
Public Function DataStartColumn() As String
    DataStartColumn = NumberToLetter(Me.DataStartColumnNumber)
End Function

Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property
'商品情報が記載されているの最後の列番号
Public Function DataEndColumnNumber() As Integer
    DataEndColumnNumber = Me.Worksheet.Cells(Me.DataStartRowNumber, Me.Worksheet.Columns.Count).End(xlToLeft).Column
End Function
'商品情報が記載されているの最後の列
Public Function DataEndColumn() As String
    DataEndColumn = NumberToLetter(Me.DataEndColumnNumber)
End Function
'次に入力する商品情報が空白の行番号
Public Function DataNextRowNumber() As Long
    DataNextRowNumber = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.ProductCodeColumnNumber).End(xlUp).Row + 1
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
'発注入力にあるデータの範囲(ヘッダー込み)
Property Get DataRange() As range
    Dim strRng As String
    strRng = Me.DataStartColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.DataEndColumn & _
             Me.DataEndRowNumber
    Set DataRange = Me.Worksheet.range(strRng)
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

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = GetRangeValue(Me.QtyRange)
End Property

'発注入力に既に入力されているデータの取得(ヘッダー込み)
Property Get data() As Variant '多次元配列で返す
    data = Me.DataRange.value
End Property
'発注入力に既に入力されているデータの取得(ヘッダーなし)
Property Get dataNoHeader() As Variant '多次元配列で返す
    dataNoHeader = RemoveFirstRow(Me.data)
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

Public Sub WriteAllData(data As Variant)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.DataStartColumnNumber, data
End Sub

'保存ファイルパス
Property Get SaveFilePath() As String
    SaveFilePath = OrderDataDirPath & "\" & SaveFileName
End Property

Public Sub OpenWorkBook()
    Set Me.Workbook = Workbooks.Open(SaveFilePath)
End Sub

Public Sub Save()
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    Me.Workbook.SaveAs Filename:=Me.SaveFilePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkBook()
    Me.Workbook.Close SaveChanges:=True
End Sub


