Attribute VB_Name = "SheetAccesser"
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
Public Const OrderWb_InputBumonCDRange As String = "A2"
Public Const OrderWb_InputUserCDRange As String = "B2"
Public Const OrderWb_OutputUserCDRange As String = "C2"
Public Const OrderWb_InputDateRange As String = "D2"

'部門コードの取得
Function GetBumonCD() As Integer

    '値の取得
    Dim Value As Integer
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputBumonCDRange).Value
    
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
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputUserCDRange).Value
    
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
    Set Value = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputDateRange).Value
    
    GetDate = Value
End Function


