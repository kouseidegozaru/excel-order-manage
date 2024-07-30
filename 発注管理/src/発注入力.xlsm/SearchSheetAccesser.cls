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

    WorkSheetName_ = "商品マスターフォーマット"
    
    ProductCodeColumnNumber_ = 3
    DataStartRowNumber_ = 2
    DataStartColumnNumber_ = 2
    CheckStateColumnNumber_ = 19
    CheckBoxColumnNumber_ = 1
    
End Sub

Private Sub NewWorkbook()
    ' 新規ワークブックの作成
    Me.Workbook = Application.Workbooks.add
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
    ProductCodeColumnNumber = ProductCDColumnNumber_
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
    
    columnNumber = Me.CheckBoxColumnNumber
    stateColumnNumber = Me.CheckStateColumnNumber
    
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    
    ' チェックボックスの追加
    Dim cb As CheckBox
    Set cb = Me.Worksheet.checkBoxes.add(Left:=ws.Cells(rowIndex, columnIndex).Left, _
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
    
    For i = 1 To ws.Cells(ws.Rows.Count, productsColumnIndex).End(xlUp).row
        If ws.Cells(i, checkedStateColumnIndex).value = True Then
            checkedProductsCode.add ws.Cells(i, productsColumnIndex).value
        End If
    Next i
    
    Set GetCheckedValue = checkedProductsCode
End Function

Property Let Cells(rowIndex As Long, columnIndex As Integer, value As Variant)
    Me.Worksheet.Cells(rowIndex, columnIndex) = value
End Property

