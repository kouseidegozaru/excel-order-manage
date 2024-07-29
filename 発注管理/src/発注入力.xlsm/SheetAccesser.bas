Attribute VB_Name = "SheetAccesser"
'定数をここで一括定義
'定数が多いためカプセル化が困難
'発注入力シートに関するデータの読み書きは基本的にここを通す

'発注データフォルダパス
Public Const OrderDataDirPath As String = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\発注管理\data"

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
Public Const OrderWb_ProductQtyColumnNumber As Integer = 10
Public Const OrderWb_ProductQtyColumn As String = "J"
Public Const OrderWb_ProductCodeRowNumber As Integer = 5
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_OutputBumonNameRange As String = "B2"
Public Const OrderWb_InputUserCDRange As String = "C2"
Public Const OrderWb_OutputUserNameRange As String = "D2"
Public Const OrderWb_InputDateRange As String = "E2"

Public Const OrderWb_IgnoreStateRange As String = "F1:F1"

'保存データ設定
Public Const DataWb_SheetName As String = "Sheet1"
Public Const DataWb_ProductCodeColumnNumber As Integer = 1
Public Const DataWb_ProductQtyColumnNumber As Integer = 2
Public Const DataWb_ProductCodeColumn As String = "A"
Public Const DataWb_ProductQtyColumn As String = "B"
Public Const DataWb_ProductCodeRowNumber As Integer = 1


'発注入力にある商品コードの範囲
Public Function OrderWb_InputProductsRange() As String
    
    OrderWb_InputProductsRange = OrderWb_ProductCodeColumn & _
                                 OrderWb_ProductCodeRowNumber & _
                                 ":" & _
                                 OrderWb_ProductCodeColumn & _
                                 OrderWb_LastProductsRow
End Function
'発注入力にある商品の数量の範囲
Public Function OrderWb_InpuQtyRange() As String
    OrderWb_InpuQtyRange = OrderWb_ProductQtyColumn & _
                            OrderWb_ProductCodeRowNumber & _
                            ":" & _
                            OrderWb_ProductQtyColumn & _
                            OrderWb_LastProductsRow
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

'商品情報が記載されている最後の行番号
Public Function OrderWb_LastProductsRow() As Long
    Dim lastRow As Long
    '発注入力に商品データがない場合行範囲をずらす(無いと範囲にヘッダー行も含まれてしまう)
    lastRow = OrderWb_NextProductsRow - 1
    If lastRow < OrderWb_ProductCodeRowNumber Then
        lastRow = OrderWb_ProductCodeRowNumber
    End If
    OrderWb_LastProductsRow = lastRow
End Function

'部門コードの取得
Function GetBumonCD() As Integer

    '値の取得
    Dim value As Integer
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputBumonCDRange).value
    
    '部門コードが存在するか
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsBumon(value) Then
        GetBumonCD = value
    Else
        GetBumonCD = 0
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
    
End Function

'担当者コードの取得
Function GetUserCD() As Integer
    
    '値の取得
    Dim value As Integer
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputUserCDRange).value
    
     '担当者コードが存在するか
    Dim DataStorage As New dataAccesser
    If DataStorage.ExistsUser(value) Then
        GetUserCD = value
    Else
        GetUserCD = 0
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
    
End Function

'対象日付の取得
Function GetDate() As Date
    Dim value As Date
    value = ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_InputDateRange).value
    
    GetDate = value
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


'担当者名の表示
Sub SetUserName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputUserNameRange).value = name
End Sub

'部門名の表示
Sub SetBumonName(name As String)
    ThisWorkbook.Sheets(OrderWb_SheetName).range(OrderWb_OutputBumonNameRange).value = name
End Sub

'保存ファイル名
Function GetSaveFileName() As String
    GetSaveFileName = "b" & GetBumonCD & "-" & _
                  "u" & GetUserCD & "-" & _
                  "d" & Format(GetDate, "yyyymmdd") & "-" & _
                  ".xlsx"
End Function

'保存ファイルパス
Function GetSaveFilePath() As String
    GetSaveFilePath = OrderDataDirPath & "\" & GetSaveFileName
End Function

'データエクセルファイルの読み込み
Public Function DataWb() As Workbook
    Set DataWb = Workbooks.Open(GetSaveFilePath)
End Function

'商品情報が記載されている最後の行番号
Public Function DataWb_LastProductsRow() As Long
    Dim wb As Workbook
    Set wb = DataWb
    
    Dim ws As Worksheet
    Dim columnNumber As Long
    
    ' 対象のシートを設定
    Set ws = wb.Sheets(DataWb_SheetName)
    columnNumber = OrderWb_ProductCodeColumnNumber
    
    ' 対象の列で最後の行を取得
    DataWb_LastProductsRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).row
End Function

'発注入力にある商品コードの範囲
Public Function DataWb_ProductsRange() As String
    
    DataWb_ProductsRange = DataWb_ProductCodeColumn & _
                            DataWb_ProductCodeRowNumber & _
                            ":" & _
                            DataWb_ProductCodeColumn & _
                            DataWb_LastProductsRow
    aa = 1
End Function
'発注入力にある商品の数量の範囲
Public Function DataWb_QtyRange() As String
    DataWb_QtyRange = DataWb_ProductQtyColumn & _
                      DataWb_ProductCodeRowNumber & _
                      ":" & _
                      DataWb_ProductQtyColumn & _
                      DataWb_LastProductsRow
End Function



