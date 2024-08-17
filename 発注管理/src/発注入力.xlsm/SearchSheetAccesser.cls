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

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer
Private CheckStateColumnIndex_ As Integer
Private CheckBoxColumnIndex_ As Integer

Private Sub Class_Initialize()

    WorkSheetName_ = "商品マスターシート"
    
    ProductCodeColumnIndex_ = 3
    CheckStateColumnIndex_ = 21
    CheckBoxColumnIndex_ = 1
        
    Set super = New SheetAccesser
    super.InitStatus DataHeaderRow:=2, _
                     DataStartRow:=3, _
                     DataStartColumn:=2
                     
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

Property Get CheckStateColumnIndex() As Integer
    CheckStateColumnIndex = CheckStateColumnIndex_
End Property

Property Get CheckBoxColumnIndex() As Integer
    CheckBoxColumnIndex = CheckBoxColumnIndex_
End Property

Public Sub AddCheckBox(rowIndex As Long)
    Dim columnIndex As Integer
    Dim stateColumnIndex As Integer
    
    columnIndex = Me.CheckBoxColumnIndex
    stateColumnIndex = Me.CheckStateColumnIndex
    
    Dim ws As Worksheet
    Set ws = Me.Worksheet
    
    ' フォームコントロールのチェックボックスを追加
    Dim cb As Shape
    Set cb = ws.Shapes.AddFormControl(xlCheckBox, _
                                      Left:=ws.Cells(rowIndex, columnIndex).Left, _
                                      Top:=ws.Cells(rowIndex, columnIndex).Top, _
                                      Width:=ws.Cells(rowIndex, columnIndex).Width, _
                                      Height:=ws.Cells(rowIndex, columnIndex).Height)
    
    ' チェックボックスのキャプションを空に設定
    cb.TextFrame.Characters.Text = ""
    
    ' チェックボックスのリンクするセルを設定
    cb.ControlFormat.LinkedCell = ws.Cells(rowIndex, stateColumnIndex).Address
End Sub

Public Function GetCheckedProductsCode() As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Me.Workbook
    Set ws = Me.Worksheet
    
    Dim productsColumn As Integer
    Dim checkedStateColumn As Integer
    
    Dim checkedProductsCode As New Collection
    Dim i As Long
    
    productsColumn = Me.ProductCodeColumnIndex
    checkedStateColumn = Me.CheckStateColumnIndex
    
    For i = 1 To ws.Cells(ws.Rows.count, productsColumn).End(xlUp).row
        If ws.Cells(i, checkedStateColumn).value = True Then
            checkedProductsCode.Add ws.Cells(i, productsColumn).value
        End If
    Next i
    
    Set GetCheckedProductsCode = checkedProductsCode
End Function

'商品コードがあるデータの範囲
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.ProductCodeColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.ProductCodeColumnIndex)
End Property
Public Sub Clear()
    Dim clearRange As Range
    Set clearRange = Me.ProductsCodeRange
    
    If WorksheetFunction.CountA(clearRange) > 0 Then
        Me.DeleteAllCheckboxes
        clearRange.EntireRow.Delete
    End If
End Sub

Sub DeleteAllCheckboxes()
    Dim ws As Worksheet
    Dim cb As OLEObject
    Dim shp As Shape

    Set ws = Me.Worksheet

    ' Form Controlsのチェックボックスを削除
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlCheckBox Then
                shp.Delete
            End If
        End If
    Next shp

End Sub

Sub ClearCheckBoxes()
    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = Me.Worksheet
    
    ' ワークシート上のすべてのシェイプをループ
    For Each shp In ws.Shapes
        ' シェイプがフォームコントロールのチェックボックスであるか確認
        If shp.Type = msoFormControl And shp.FormControlType = xlCheckBox Then
            ' チェックボックスがオンの場合
            If shp.ControlFormat.value = xlOn Then
                shp.ControlFormat.value = xlOff
            End If
        End If
    Next shp
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
