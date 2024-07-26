Attribute VB_Name = "SheetAccesser"
'定数をここで一括定義
'定数が多いためカプセル化が困難

'商品検索フォームのセル番地の設定
Public Const SearchWb_SheetName As String = "商品マスターフォーマット"
Public Const SearchWb_StateColumnNumber As Integer = 19
Public Const SearchWb_CheckBoxColumnNumber As Integer = 1
Public Const SearchWb_DataStartColumnNumber As Integer = 2
Public Const SearchWb_DataStartRowNumber As Integer = 2
Public Const SearchWb_ProductCodeColumnNumber As Integer = 3

'発注入力フォームのセル番地の設定
Public Const OrderWb_SheetName As String = "発注入力"
Public Const OrderWb_ProductCodeColumnNumber As Integer = 1
Public Const OrderWb_ProductCodeColumn As String = "A"
Public Const OrderWb_ProductCodeRowNumber As Integer = 5
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_OutputBumonNameRange As String = "B2"
Public Const OrderWb_InputUserCDRange As String = "C2"
Public Const OrderWb_OutputUserNameRange As String = "D2"
Public Const OrderWb_InputDateRange As String = "E2"
Public Function OrderWb_InputProductsRange() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim columnNumber As Long
    
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = OrderWb_NextProductsRow - 1
    If lastRow < OrderWb_ProductCodeRowNumber Then
        lastRow = OrderWb_ProductCodeRowNumber
    End If
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    OrderWb_InputProductsRange = OrderWb_ProductCodeColumn & _
                                 OrderWb_ProductCodeRowNumber & _
                                 ":" & _
                                 OrderWb_ProductCodeColumn & _
                                 lastRow
End Function
'次に入力する商品情報が空白の行番号
Public Function OrderWb_NextProductsRow() As Long
    Dim ws As Worksheet
    Dim columnNumber As Long
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    ' 対象の列で最後の行を取得
    OrderWb_NextProductsRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row + 1
End Function

'部門コードの取得
Function GetBumonCD() As Integer

    '値の取得
    Dim Value As Integer
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputBumonCDRange).Value
    
    '部門コードが存在するか
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsBumon(Value) Then
        GetBumonCD = Value
    Else
        GetBumonCD = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
    
End Function

'担当者コードの取得
Function GetUserCD() As Integer
    
    '値の取得
    Dim Value As Integer
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputUserCDRange).Value
    
     '担当者コードが存在するか
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsUser(Value) Then
        GetUserCD = Value
    Else
        GetUserCD = 0
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
    
End Function

'対象日付の取得
Function GetDate() As Date
    Dim Value As Date
    Value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputDateRange).Value
    
    GetDate = Value
End Function
'発注入力に既に入力されている商品コードの取得
Function GetProductsCD() As Collection
    ' 対象のシートを設定
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OrderWb_SheetName)
    
    Dim rng As range
    Set rng = ws.range(OrderWb_InputProductsRange)
    
    Set GetProductsCD = GetRangeValue(rng)
End Function

Private Function GetRangeValue(rng As range) As Collection
    Dim cell As range
    Dim col As New Collection
    
    ' 範囲内の各セルをループ
    For Each cell In rng
        ' 空白でないセルの場合、Collectionに追加
        If cell.Value <> "" Then
            col.Add cell.Value
        End If
    Next cell
    
    Set GetRangeValue = col
    
End Function

'担当者名の表示
Sub SetUserName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputUserNameRange).Value = name
End Sub

'部門名の表示
Sub SetBumonName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputBumonNameRange).Value = name
End Sub


