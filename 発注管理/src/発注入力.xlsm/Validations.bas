Attribute VB_Name = "Validations"
Sub SetValidations()
    SetBumonCD
    SetUserCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputBumonCDRange)
    
    With rng.Validation
        .Delete ' 既存のバリデーションを削除
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' 数値型のバリデーションを追加
        .IgnoreBlank = True ' 空白セルを無視
        .InCellDropdown = True ' ドロップダウンリストを表示
        .InputTitle = "部門コード"
        .ErrorTitle = "入力エラー"
        .InputMessage = "数値を入力してください。"
        .ErrorMessage = "入力値が数値ではありません。"
        .ShowInput = True ' 入力メッセージを表示
        .ShowError = True ' エラーメッセージを表示
    End With
End Sub

Private Sub SetUserCD()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputUserCDRange)
    
    With rng.Validation
        .Delete ' 既存のバリデーションを削除
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10000 ' 数値型のバリデーションを追加
        .IgnoreBlank = True ' 空白セルを無視
        .InCellDropdown = True ' ドロップダウンリストを表示
        .InputTitle = "担当者コード"
        .ErrorTitle = "入力エラー"
        .InputMessage = "数値を入力してください。"
        .ErrorMessage = "入力値が数値ではありません。"
        .ShowInput = True ' 入力メッセージを表示
        .ShowError = True ' エラーメッセージを表示
    End With
End Sub

Private Sub SetDate()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets(OrderWb_SheetName).Range(OrderWb_InputDateRange)
    
    With rng.Validation
        .Delete ' 既存のバリデーションを削除
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1/1/1900", Formula2:="12/31/2100" ' 日付型のバリデーションを追加
        .IgnoreBlank = True ' 空白セルを無視
        .InCellDropdown = True ' ドロップダウンリストを表示
        .InputTitle = "発注日付"
        .ErrorTitle = "入力エラー"
        .InputMessage = "有効な日付を入力してください。"
        .ErrorMessage = "入力値が有効な日付ではありません。"
        .ShowInput = True ' 入力メッセージを表示
        .ShowError = True ' エラーメッセージを表示
    End With
End Sub
