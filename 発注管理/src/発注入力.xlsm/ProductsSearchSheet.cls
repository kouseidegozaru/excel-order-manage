VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductsSearchSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'商品検索シートの表示処理クラス

Private rs As ADODB.recordSet
Private wb As Workbook
Private ws As Worksheet
Private checkBoxes As Collection

Private stateColumnNumber As Integer
Private checkBoxColumnNumber As Integer
Private dataStartColumnNumber As Integer
Private dataStartRowNumber As Long


Public Sub Initialize(recordSet As ADODB.recordSet)

    Set rs = recordSet
    Call newWorkbook
    Set checkBoxes = New Collection
    
    stateColumnNumber = SearchWb_StateColumnNumber
    checkBoxColumnNumber = SearchWb_CheckBoxColumnNumber
    dataStartColumnNumber = SearchWb_DataStartColumnNumber
    dataStartRowNumber = SearchWb_DataStartRowNumber
End Sub
Private Sub newWorkbook()
    ' コピー元のワークブックを選択する
    Dim sourceWb As Workbook
    Set sourceWb = ThisWorkbook ' 既存のワークブック
    
    ' 新規ワークブックの作成
    Set wb = Application.Workbooks.Add
    '最大化
    wb.Windows(1).WindowState = xlMaximized

    ' コピー元のワークシートを選択する
    Dim sourceWs As Worksheet
    Set sourceWs = sourceWb.Sheets(SearchWb_SheetName)

    ' ワークシートをコピーする
    sourceWs.Copy Before:=wb.Sheets(1)

    Set ws = wb.Sheets(1) ' 新しいワークブックの最初のワークシートを選択する
End Sub
Public Sub ExportRecordSet()
    Dim i As Long, j As Long
    Dim fld As Object
    Dim row As Long: row = dataStartRowNumber
    
    ' ヘッダーの書き込み
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(row, i + dataStartColumnNumber).value = rs.Fields(i).name
    Next i
    
    ' データの書き込み
    rs.MoveFirst
    Do While Not rs.EOF
        row = row + 1
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(row, i + dataStartColumnNumber).value = rs.Fields(i).value
        Next i
        
        ' チェックボックスの追加
        Call AddCheckBox(row, checkBoxColumnNumber, stateColumnNumber)
        
        rs.MoveNext
    Loop
End Sub


Private Sub AddCheckBox(RowNumber As Long, columnNumber As Integer, stateColumnNumber As Integer)
    ' チェックボックスの追加
    Dim cb As CheckBox
    Set cb = ws.checkBoxes.Add(Left:=ws.Cells(RowNumber, columnNumber).Left, _
                               Top:=ws.Cells(RowNumber, columnNumber).Top, _
                               Width:=ws.Cells(RowNumber, columnNumber).Width, _
                               Height:=ws.Cells(RowNumber, columnNumber).Height)
    cb.Caption = ""
    cb.LinkedCell = ws.Cells(RowNumber, stateColumnNumber).Address
End Sub

