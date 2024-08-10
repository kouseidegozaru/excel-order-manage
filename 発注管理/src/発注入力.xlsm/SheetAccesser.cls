VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'シートにアクセスするための抽象クラス

Private WorkBook_ As Workbook
Private WorkSheet_ As Worksheet

Private DataHeaderRowIndex_ As Long
Private DataStartRowIndex_ As Long
Private DataStartColumnIndex_ As Integer

Public Sub InitStatus(Optional DataHeaderRow As Long = 1, Optional DataStartRow As Long = 1, Optional DataStartColumn As Integer = 1)
    DataHeaderRowIndex_ = DataHeaderRow
    DataStartRowIndex_ = DataStartRow
    DataStartColumnIndex_ = DataStartColumn
End Sub
Public Sub InitWorkBook(wb As Workbook)
    Set WorkBook_ = wb
End Sub

Public Sub InitWorkSheet(ws As Worksheet)
    Set WorkSheet_ = ws
End Sub

Property Get Workbook() As Workbook
    Set Workbook = WorkBook_
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = WorkSheet_
End Property

Property Get DataHeaderRowIndex() As Long
    DataHeaderRowIndex = DataHeaderRowIndex_
End Property

Property Get DataStartRowIndex() As Long
    DataStartRowIndex = DataStartRowIndex_
End Property

Property Get DataStartColumnIndex() As Integer
    DataStartColumnIndex = DataStartColumnIndex_
End Property
'商品情報が記載されているの最後の列番号
Property Get DataEndColumnIndex() As Integer
    DataEndColumnIndex = Me.Worksheet.Cells(Me.DataHeaderRowIndex, Me.Worksheet.Columns.Count).End(xlToLeft).Column
End Property
'次に入力する商品情報が空白の行番号
Property Get DataNextRowIndex() As Long
    DataNextRowIndex = Me.Worksheet.Cells(Me.Worksheet.Rows.Count, Me.DataStartColumnIndex).End(xlUp).Row + 1
End Property

'商品情報が記載されている最後の行番号
Property Get DataEndRowIndex() As Long
    Dim lastRow As Long
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = Me.DataNextRowIndex - 1
    If lastRow < Me.DataStartRowIndex Then
        lastRow = Me.DataStartRowIndex
    End If
    DataEndRowIndex = lastRow
End Property

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex).value = value
End Property

Public Sub WriteAllData(targetData As Variant) '二次元配列を受け取る
    WriteMatrixArray Me.Worksheet, Me.DataStartRowIndex, Me.ProductCodeColumnIndex, targetData
End Sub

Property Get AllDataRange() As Range
    GetAllData = GetRange(Me.Worksheet, Me.DataStartRowIndex, Me.DataStartColumnIndex, Me.DataEndRowIndex, Me.DataEndColumnIndex).value
End Property

Public Function GetAllData() As Variant
    GetAllData = Me.AllDataRange.value
End Function

Public Function GetAllDara_NoHead() As Variant
    GetAllDara_NoHead = RemoveFirstRow(Me.GetAllData)
End Function

'''汎用メソッド'''
Public Function GetRange(startRow As Long, startCol As Long, endRow As Long, endCol As Long) As Range
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    Set GetRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
End Function

' データをシートに書き込む
Public Sub WriteMatrixArray(startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCount As Long, colCount As Long
    
    Dim ws As Worksheet
    Set ws = Me.Worksheet

    ' writeDataが配列か確認
    If IsArray(writeData) Then
        rowCount = 0
        ' 配列の場合
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            colCount = 0
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + colCount, startColIndex + rowCount).value = writeData(i, j)
                colCount = colCount + 1
            Next j
            rowCount = rowCount + 1
        Next i
    Else
        ' エラーハンドリング
        Err.Raise vbObjectError + 9999, "writeData", "writeDataは配列でなければなりません。"
    End If
End Sub

'rangeで指定した範囲が一行または一列の場合に一次元の配列に格納する
Public Function RangeToOneDimArray(rng As Range) As Variant
    Dim arr As Variant
    Dim oneDimArr() As Variant
    Dim i As Integer

    arr = rng.value

    ' 一行か一列かを判定
    If rng.Rows.Count = 1 Then
        ' 一行の場合
        ReDim oneDimArr(1 To rng.Columns.Count)
        For i = 1 To rng.Columns.Count
            oneDimArr(i) = arr(1, i)
        Next i
    ElseIf rng.Columns.Count = 1 Then
        ' 一列の場合
        ReDim oneDimArr(1 To rng.Rows.Count)
        For i = 1 To rng.Rows.Count
            oneDimArr(i) = arr(i, 1)
        Next i
    End If

    RangeToOneDimArray = oneDimArr
End Function

