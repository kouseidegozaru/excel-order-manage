VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadSheetAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'発注確認シートへのアクセサ

'SheetAccesserクラスを疑似的に継承
Private super As SheetAccesser

Private WorkSheetName_ As String

Private ProductCodeColumnIndex_ As Integer

Private CheckStateColumnIndex_ As Integer
Private CheckBoxColumnIndex_ As Integer

Private InputBumonCodeRange_ As String
Private OutputBumonNameRange_ As String
Private InputTargetDateRange_ As String

Private FilePropertyStartColumnIndex_ As Integer
Private FilePropertyUserNameRowIndex_ As Long
Private FilePropertyCompleteRowIndex_ As Long
Private FilePropertyUpdateTimeRowIndex_ As Long

Private CompleteString_ As String

Private Sub Class_Initialize()

    'シート名
    WorkSheetName_ = "発注確認"
    
    '発注データを入力済みの表示するユーザーフラグ
    CompleteString_ = "◎"
    
    '商品コードの列番号
    ProductCodeColumnIndex_ = 2
    'チェックボックスの状態の列番号
    CheckStateColumnIndex_ = 19
    'チェックボックスの列番号
    CheckBoxColumnIndex_ = 1
    
    'ファイル情報
    '開始列
    FilePropertyStartColumnIndex_ = 2
    '担当者名表示行
    FilePropertyUserNameRowIndex_ = 4
    '発注入力フラグ表示行
    FilePropertyCompleteRowIndex_ = 5
    '最終更新日表示行
    FilePropertyUpdateTimeRowIndex_ = 6
    
    '部門コード入力セル
    InputBumonCodeRange_ = "A2"
    '部門名表示セル
    OutputBumonNameRange_ = "B2"
    '発注日表示セル
    InputTargetDateRange_ = "C2"
    
    '疑似抽象クラスをインスタンス化
    Set super = New SheetAccesser
    'データのヘッダー行、開始行、開始列を指定
    super.InitStatus DataHeaderRow:=8, _
                     DataStartRow:=9, _
                     DataStartColumn:=2
    
    InitWorkbook
    InitWorkSheet
End Sub

'ワークブックを設定
Private Sub InitWorkbook()
    super.InitWorkbook ThisWorkbook
End Sub

'ワークシートを設定
Private Sub InitWorkSheet()
    super.InitWorkSheet super.Workbook.Sheets(WorkSheetName_)
End Sub

'商品コードの列番号
Property Get ProductCodeColumnIndex() As Integer
    ProductCodeColumnIndex = ProductCodeColumnIndex_
End Property
'チェックボックスの状態の列番号
Property Get CheckStateColumnIndex() As Integer
    CheckStateColumnIndex = CheckStateColumnIndex_
End Property
'チェックボックスの列番号
Property Get CheckBoxColumnIndex() As Integer
    CheckBoxColumnIndex = CheckBoxColumnIndex_
End Property
'ファイル情報の開始列
Property Get FilePropertyStartColumnIndex() As Integer
    FilePropertyStartColumnIndex = FilePropertyStartColumnIndex_
End Property

'次に入力するファイル情報が空白の列番号
Public Function FilePropertyNextColumnIndex() As Integer
    FilePropertyNextColumnIndex = Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, Me.Worksheet.Columns.Count).End(xlToLeft).Column + 1
End Function

'ファイル情報が記載されている最後の列番号
Property Get FilePropertyEndColumnIndex() As Long
    Dim lastColumn As Long
    'ファイル情報にデータがない場合列範囲をずらす(無いと範囲にヘッダー列も含まれてしまう)
    lastColumn = Me.FilePropertyNextColumnIndex - 1
    If lastColumn < Me.FilePropertyStartColumnIndex Then
        lastColumn = Me.FilePropertyStartColumnIndex
    End If
    FilePropertyEndColumnIndex = lastColumn
End Property

'ファイル情報の範囲
Property Get FilePropertyRange() As Range
    Set FilePropertyRange = super.GetRange(FilePropertyUserNameRowIndex_, _
                                           Me.FilePropertyStartColumnIndex, _
                                           FilePropertyUpdateTimeRowIndex_, _
                                           Me.FilePropertyEndColumnIndex)
End Property

'ファイル情報をクリア
Public Sub ClearFileProperty()
    Me.FilePropertyRange.Delete Shift:=xlToLeft
End Sub

'ファイル情報を追加
Public Sub AddFileProperty(userName As String, IsCompleted As Boolean, updatedTime As Date)
    
    '入力対象列
    Dim lastColumn As Integer
    lastColumn = FilePropertyNextColumnIndex
    
    '発注入力フラグ文字列
    Dim writeCompleteString As String
    writeCompleteString = ""
    
    '発注入力をした担当者の場合
    If IsCompleted Then
        writeCompleteString = CompleteString_
        
        '更新日時書き込み
        Me.Worksheet.Cells(FilePropertyUpdateTimeRowIndex_, lastColumn) = updatedTime
    End If
    
    '担当者名表示
    Me.Worksheet.Cells(FilePropertyUserNameRowIndex_, lastColumn) = userName
    '発注入力フラグ表示
    Me.Worksheet.Cells(FilePropertyCompleteRowIndex_, lastColumn) = writeCompleteString
End Sub

'全てのチェックボックスの削除
Sub DeleteAllCheckboxes()
    Dim ws As Worksheet
    Dim cb As OLEObject
    Dim shp As Shape

    Set ws = Me.Worksheet

    ' Form Controlsのチェックボックスを削除
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            'コントロールがチェックボックスの場合
            If shp.FormControlType = xlCheckBox Then
                '削除
                shp.Delete
            End If
        End If
    Next shp

End Sub

'表示されている発注データの削除
Public Sub ClearData()
    Dim clearRange As Range
    Set clearRange = Me.ProductsCodeRange
    '行が存在する場合
    If WorksheetFunction.CountA(clearRange) > 0 Then
        'チェックボックスの削除
        Me.DeleteAllCheckboxes
        '商品情報の削除
        clearRange.EntireRow.Delete
    End If
End Sub

'発注入力にある商品コードの範囲
Property Get ProductsCodeRange() As Range
    Set ProductsCodeRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.ProductCodeColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.ProductCodeColumnIndex)
End Property

'発注入力にあるチェックボックス状態の範囲
Property Get CheckStateRange() As Range
    Set CheckStateRange = super.GetRange(Me.DataStartRowIndex, _
                                           Me.CheckStateColumnIndex, _
                                           Me.DataEndRowIndex, _
                                           Me.CheckStateColumnIndex)
End Property

'部門コードの取得
Property Get bumonCode() As Integer
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputBumonCodeRange_).value
    
    'バリデーションチェックを行い失敗したら終了
    CheckExistsBumon value
    
    bumonCode = CInt(value)

End Property

'対象日付の取得
Property Get targetDate() As Date
    '値の取得
    Dim value As Variant
    value = Me.Worksheet.Range(InputTargetDateRange_).value
    
    'バリデーションチェックを行い失敗したら終了
    CheckDateFormat value
    
    targetDate = CDate(value)
End Property

'部門コードの入力セル
Property Get BumonCodeRange() As Range
    Set BumonCodeRange = Me.Worksheet.Range(InputBumonCodeRange_)
End Property
'発注日の入力セル
Property Get TargetDateRange() As Range
    Set TargetDateRange = Me.Worksheet.Range(InputTargetDateRange_)
End Property

'部門名の表示
Property Let BumonName(name As String)
    Me.Worksheet.Range(OutputBumonNameRange_).value = name
End Property

'シートに入力されている各項目ごとにグループ化して数量と仕入金額を合算
Public Function AllGroupData() As ADODB.Recordset
    Dim dataAccessor As New SqlExecutor
    Dim rs As ADODB.Recordset
    Dim FilePath As String
    Dim sheetName As String
    Dim rangeName As String
    Dim query As String
    
    ' Excel情報
    FilePath = Me.Workbook.FullName
    sheetName = Me.Worksheet.name
    rangeName = Replace(Me.AllDataRange.Address(False, False), "$", "")
    
    
    ' 実行するクエリ
    query = "SELECT 品CD, 入荷時間帯, 仕入先名, 発注先名, メーカー名, 品名, 規格, 計区, 入数, 合わせ数, SUM(数量) AS 合計数量, 荷姿, 仕入単価, SUM(仕入金額) AS 合計仕入金額, LT, 商品備考, メーカー読み " & _
           "FROM [" & sheetName & "$" & rangeName & "] " & _
           "GROUP BY 品CD, 入荷時間帯, 仕入先名, 発注先名, メーカー名, 品名, 規格, 計区, 入数, 合わせ数, 荷姿, 仕入単価, LT, 商品備考, メーカー読み"
    
    ' Excelファイルにクエリを実行
    Set rs = dataAccessor.ExecuteExcelQuery(FilePath, query)
    
    Set AllGroupData = rs
End Function

'対象の行のチェックボックスのステータス変更
Property Let CheckBoxState(rowIndex As Long, value As Boolean)
    Me.Cells(rowIndex, Me.CheckStateColumnIndex) = value
End Property

'対象の商品のチェックボックスをオンにする
Public Sub OrderedIsTrue(targetProductsCode As Variant)
    
    'データの終了行と開始行
    Dim startRowIndex As Long
    startRowIndex = Me.DataStartRowIndex
    Dim endRowIndex As Long
    endRowIndex = Me.DataEndRowIndex
    
    '現在表示されている商品コード
    Dim displayProductsCode As Variant
    
    For i = startRowIndex To endRowIndex
        '商品コード
        displayProductsCode = Me.Worksheet.Cells(i, Me.ProductCodeColumnIndex).value
        '対象の商品コードの場合チェックボックスをオン
        If targetProductsCode = displayProductsCode Then
            Me.CheckBoxState(CInt(i)) = True
        End If
    Next i
    
End Sub

'行を指定してチェックボックスを追加
Public Function AddCheckBox(rowIndex As Long) As Shape
    Dim columnIndex As Integer
    Dim stateColumnIndex As Integer
    
    'チェックボックスの列
    columnIndex = Me.CheckBoxColumnIndex
    'チェックボックスの状態の列
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
    
    '追加したチェックボックスオブジェクトを返却
    Set AddCheckBox = cb
End Function

'チェックされた商品コードの取得
Public Function GetCheckedProductsCode() As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Me.Workbook
    Set ws = Me.Worksheet
    
    '商品コード列
    Dim productsColumn As Integer
    'チェックボックスの状態の列
    Dim checkedStateColumn As Integer
    
    'チェックされた商品コードの格納用
    Dim checkedProductsCode As New Collection
    Dim i As Long
    
    productsColumn = Me.ProductCodeColumnIndex
    checkedStateColumn = Me.CheckStateColumnIndex
    
    '商品コードが存在する行番号でループ
    For i = 1 To ws.Cells(ws.Rows.Count, productsColumn).End(xlUp).row
        'チェックされていた場合
        If ws.Cells(i, checkedStateColumn).value = True Then
            '商品コードを追加
            checkedProductsCode.Add ws.Cells(i, productsColumn).value
        End If
    Next i
    
    Set GetCheckedProductsCode = checkedProductsCode
End Function

'条件付き書式を設定
Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim startCol As Integer
    Dim endCol As Integer
    Dim checkCol As Integer
    Dim rng As Range
    Dim formatCondition As formatCondition
    
    ' 対象のシートを設定
    Set ws = Me.Worksheet
    
    ' 列の指定
    startCol = Me.DataStartColumnIndex
    endCol = Me.DataEndColumnIndex
    checkCol = Me.CheckStateColumnIndex
    
    ' 行の指定
    lastRow = Me.DataEndRowIndex
    startRow = Me.DataStartRowIndex
    
    ' 各行ごとに条件付き書式を設定
    For i = startRow To lastRow
        ' S列がTrueの場合、AからS列の範囲に条件付き書式を設定
        Set rng = ws.Range(ws.Cells(i, startCol), ws.Cells(i, endCol))
        
        ' 条件付き書式をクリア（既存の書式を削除）
        rng.FormatConditions.Delete
        
        ' 条件付き書式を追加
        Set formatCondition = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ws.Cells(i, checkCol).Address & "=TRUE")
        
        ' 背景色を設定（黄色）
        formatCondition.Interior.Color = RGB(255, 255, 0) ' 黄色
    Next i
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

