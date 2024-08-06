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

Private fileName_ As String

Private ProductCodeColumnNumber_ As Integer
Private QtyColumnNumber_ As Integer
Private DataStartRowNumber_ As Long

Private Sub Class_Initialize()

    OrderDataDirPath = "\\mfh008-mfre\FS事業部\新）発注ノート\発注ノートシステム\data"

    WorkSheetName_ = "Sheet1"
    
    ProductCodeColumnNumber_ = 1
    QtyColumnNumber_ = 2
    DataStartRowNumber_ = 1
    
End Sub

Property Get DataDirPath()
    DataDirPath = OrderDataDirPath
End Property

Property Get filePath()
    filePath = Me.DataDirPath & "\" & fileName_
End Property

Property Let FileName(targetFileName As String)
    fileName_ = targetFileName
    
    InitWorkBook
    InitWorkSheet
End Property

Public Sub InitWorkBook()
    Set WorkBook_ = Workbooks.Open(Me.filePath)
End Sub

Public Sub InitWorkSheet()
    Set WorkSheet_ = Me.Workbook.Sheets(WorkSheetName_)
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

Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property

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
Property Get ProductsCodeRange() As Range
    Dim strRng As String
    strRng = Me.ProductCodeColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.ProductCodeColumn & _
             Me.DataEndRowNumber
    Set ProductsCodeRange = Me.Worksheet.Range(strRng)
End Property

'発注入力にある商品の数量の範囲
Property Get QtyRange() As Range
    Dim strRng As String
    strRng = Me.QtyColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.QtyColumn & _
             Me.DataEndRowNumber
    Set QtyRange = Me.Worksheet.Range(strRng)
End Property

'発注入力に既に入力されている商品コードの取得
Property Get productsCode() As Collection
    Set productsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

'発注入力に既に入力されている数量の取得
Property Get Qty() As Collection
    Set Qty = GetRangeValue(Me.QtyRange)
End Property

'商品コードと数量の辞書型データ
Public Function DataDict() As Scripting.Dictionary
    Dim ProductsCodeData As Collection
    Dim QtyData As Collection
    Dim dict As New Scripting.Dictionary
    
    Set ProductsCodeData = Me.productsCode
    Set QtyData = Me.Qty
    
    For i = 1 To ProductsCodeData.Count
        dict.Add ProductsCodeData(i), QtyData(i)
    Next
    
    Set DataDict = dict
        
End Function

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

Public Sub WriteProductsCode(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.ProductCodeColumnNumber, writeCollection
End Sub

Public Sub WriteQty(writeCollection As Collection)
    writeData Me.Worksheet, Me.DataStartRowNumber, Me.QtyColumnNumber, writeCollection
End Sub

Public Sub Save()
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    Me.Workbook.SaveAs FileName:=Me.SaveFilePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkBook()
    Me.Workbook.Close SaveChanges:=True
End Sub




