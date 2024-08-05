VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注入力フォームのセル番地の設定

Private WorkBook_ As Workbook
Private WorkSheetName_ As String
Private WorkSheet_ As Worksheet

Private ProductCodeColumnNumber_ As Integer
Private DataStartRowNumber_ As Long
Private DataStartColumnNumber_ As Integer
Private CheckStateColumnNumber_ As Integer
Private CheckBoxColumnNumber_ As Integer

Private Sub Class_Initialize()

    WorkSheetName_ = "商品マスターシート"
    
    ProductCodeColumnNumber_ = 3
    DataStartRowNumber_ = 3
    DataStartColumnNumber_ = 2
    CheckStateColumnNumber_ = 19
    CheckBoxColumnNumber_ = 1
    
End Sub

Public Sub NewWorkbook()
    ' 新規ワークブックの作成
    Set Me.Workbook = Application.Workbooks.add
End Sub

Property Set Workbook(wb As Workbook)
    Set WorkBook_ = wb
End Property

Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property

Public Sub CopyWorkSheet(baseWorkSheet As Worksheet)
    '最大化
    Me.Workbook.Windows(1).WindowState = xlMaximized
    
    ' ワークシートをコピーする
    baseWorkSheet.Copy Before:=Me.Workbook.Sheets(1)
    
    InitWorkSheet
End Sub

Public Sub InitWorkSheet()
    Set WorkSheet_ = Me.Workbook.Sheets(WorkSheetName_)
End Sub

Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
End Property

Property Get ProductCodeColumnNumber() As Integer
    ProductCodeColumnNumber = ProductCodeColumnNumber_
End Property

Property Get ProductCodeColumn() As String
    ProductCodeColumn = NumberToLetter(Me.ProductCodeColumnNumber)
End Property

Property Get DataStartRowNumber() As Long
    DataStartRowNumber = DataStartRowNumber_
End Property

Property Get DataStartColumnNumber() As Integer
    DataStartColumnNumber = DataStartColumnNumber_
End Property

Property Get CheckStateColumnNumber() As Integer
    CheckStateColumnNumber = CheckStateColumnNumber_
End Property

Property Get CheckBoxColumnNumber() As Integer
    CheckBoxColumnNumber = CheckBoxColumnNumber_
End Property

Public Sub AddCheckBox(rowIndex As Long)
    Dim columnIndex As Integer
    Dim stateColumnIndex As Integer
    
    columnIndex = Me.CheckBoxColumnNumber
    stateColumnIndex = Me.CheckStateColumnNumber
    
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    
    ' チェックボックスの追加
    Dim cb As CheckBox
    Set cb = ws.checkBoxes.add(Left:=ws.Cells(rowIndex, columnIndex).Left, _
                               Top:=ws.Cells(rowIndex, columnIndex).Top, _
                               Width:=ws.Cells(rowIndex, columnIndex).Width, _
                               Height:=ws.Cells(rowIndex, columnIndex).Height)
    cb.Caption = ""
    cb.LinkedCell = ws.Cells(rowIndex, stateColumnIndex).Address
End Sub

'検索フォームの選択された行の商品コードを取得
Public Function GetCheckedProductsCode() As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Me.Workbook
    Set ws = Me.Worksheet
    
    Dim productsColumnIndex As Integer
    Dim checkedStateColumnIndex As Integer
    
    Dim checkedProductsCode As New Collection
    Dim i As Long
    
    productsColumnIndex = Me.ProductCodeColumnNumber
    checkedStateColumnIndex = Me.CheckStateColumnNumber
    
    For i = 1 To ws.Cells(ws.Rows.Count, productsColumnIndex).End(xlUp).row
        If ws.Cells(i, checkedStateColumnIndex).value = True Then
            checkedProductsCode.add ws.Cells(i, productsColumnIndex).value
        End If
    Next i
    
    Set GetCheckedProductsCode = checkedProductsCode
End Function

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

'次に入力する商品情報が空白の行番号
Public Function DataNextRowNumber() As Long
    DataNextRowNumber = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.ProductCodeColumnNumber).End(xlUp).row + 1
End Function
'情報が記載されている最後の行番号
Property Get DataEndRowNumber() As Long
    Dim lastRow As Long
    '商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = Me.DataNextRowNumber - 1
    If lastRow < Me.DataStartRowNumber Then
        lastRow = Me.DataStartRowNumber
    End If
    DataEndRowNumber = lastRow
End Property
'商品コードがあるデータの範囲
Property Get DataRange() As range
    Dim strRng As String
    strRng = Me.ProductCodeColumn & _
             Me.DataStartRowNumber & _
             ":" & _
             Me.ProductCodeColumn & _
             Me.DataEndRowNumber
    Set DataRange = Me.Worksheet.range(strRng)
End Property
Public Sub Clear()
    DataRange.EntireRow.Delete
End Sub


