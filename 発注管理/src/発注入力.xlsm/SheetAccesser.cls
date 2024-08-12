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
Public Sub InitWorkbook(wb As Workbook)
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
    DataEndColumnIndex = Me.Worksheet.Cells(Me.DataHeaderRowIndex, Me.Worksheet.Columns.count).End(xlToLeft).Column
End Property
'次に入力する商品情報が空白の行番号
Property Get DataNextRowIndex() As Long
    DataNextRowIndex = Me.Worksheet.Cells(Me.Worksheet.Rows.count, Me.DataStartColumnIndex).End(xlUp).row + 1
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

Public Sub WriteAllData(targetData As Collection) '二次元配列を受け取る
    WriteMatrixCollection Me.DataNextRowIndex, Me.DataStartColumnIndex, targetData
End Sub

Property Get AllDataRange() As Range
    Set AllDataRange = GetRange(Me.DataHeaderRowIndex, Me.DataStartColumnIndex, Me.DataEndRowIndex, Me.DataEndColumnIndex)
End Property

Public Function GetAllData() As Collection
    Set GetAllData = ArrayToCollection(Me.AllDataRange.value)
End Function

Public Function GetAllData_NoHead() As Collection
    Set GetAllData_NoHead = RemoveFirstRow(Me.GetAllData)
End Function

'''汎用メソッド'''
Public Function GetRange(startRow As Long, startCol As Long, endRow As Long, endCol As Long) As Range
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    Set GetRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
End Function

' データをシートに書き込む
Public Sub WriteMatrixCollection(startRowIndex As Long, startColIndex As Integer, writeData As Collection)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCollection As Collection
    Dim ws As Worksheet
    Set ws = Me.Worksheet

    ' writeDataがCollectionか確認
    If TypeName(writeData) = "Collection" Then
        ' Collectionの場合
        rowCount = 0
        For Each rowCollection In writeData
            colCount = 0
            For Each item In rowCollection
                ws.Cells(startRowIndex + rowCount, startColIndex + colCount).value = item
                colCount = colCount + 1
            Next item
            rowCount = rowCount + 1
        Next rowCollection
    Else
        ' エラーハンドリング
        Err.Raise vbObjectError + 9999, "writeData", "writeDataはCollectionでなければなりません。"
    End If
End Sub

' rangeで指定した範囲が一行または一列の場合に一次元のCollectionに格納する
Public Function RangeToOneDimCollection(rng As Range) As Collection
    Dim arr As Variant
    Dim oneDimCollection As New Collection
    Dim i As Integer

    arr = rng.value
    
    If IsEmpty(arr) Then
        Set RangeToOneDimCollection = oneDimCollection
        Exit Function
    End If

    ' 一行か一列かを判定
    If rng.Rows.count = 1 Then
        ' 一行の場合
        For i = 1 To rng.Columns.count
            oneDimCollection.Add arr(1, i)
        Next i
    ElseIf rng.Columns.count = 1 Then
        ' 一列の場合
        For i = 1 To rng.Rows.count
            oneDimCollection.Add arr(i, 1)
        Next i
    End If

    Set RangeToOneDimCollection = oneDimCollection
End Function


