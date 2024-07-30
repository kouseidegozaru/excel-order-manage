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
Private DataStartRowNumber_ As Long


Private Sub Class_Initialize()

    OrderDataDirPath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\発注管理\data"

    WorkSheetName_ = "Sheet1"
    
    ProductCodeColumnNumber_ = 1
    QtyColumnNumber_ = 2
    DataStartRowNumber_ = 1
    
    InitWorkBook
    InitWorkSheet
    
End Sub
'保存ファイル名
Property Get SaveFileName() As String
    Dim order As New OrderSheetAccesser
    
    SaveFileName = "b" & order.BumonCode & "-" & _
                  "u" & order.UserCode & "-" & _
                  "d" & Format(order.TargetDate, "yyyymmdd") & "-" & _
                  ".xlsx"
End Property

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

'発注入力に既に入力されている商品コードの取得
Property Get ProductsCode() As Collection
    Set ProductsCode = GetRangeValue(Me.ProductsCodeRange)
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex) = value
End Property


'保存ファイルパス
Property Get SaveFilePath() As String
    aveFilePath = OrderDataDirPath & "\" & GetSaveFileName
End Property

'データエクセルファイルの読み込み
Public Function DataWb() As Workbook
    Set DataWb = Workbooks.Open(GetSaveFilePath)
End Function

Public Sub Save()
        
    ' 上書き保存のために警告メッセージをオフにする
    Application.DisplayAlerts = False
    
    ' ワークブックを保存（既存のファイルがあれば上書き保存）
    Me.Workbook.SaveAs Filename:=Me.SaveFilePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
    ' 警告メッセージを再度有効にする
    Application.DisplayAlerts = True
End Sub

Public Sub CloseWorkBook()
    Me.Workbook.Close
End Sub





