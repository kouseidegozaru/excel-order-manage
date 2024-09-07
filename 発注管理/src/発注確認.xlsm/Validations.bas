Attribute VB_Name = "Validations"
'セルにバリデーションチェックを設定
Sub SetValidations()
    SetBumonCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim load As New LoadSheetAccesser
    Dim rng As Range
    Set rng = load.BumonCodeRange
    
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

Private Sub SetDate()
    Dim load As New LoadSheetAccesser
    Dim rng As Range
    Set rng = load.TargetDateRange
    
    With rng.Validation
        .Delete ' 既存のバリデーションを削除
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1/1/1900", Formula2:="12/31/2100" ' 日付型のバリデーションを追加
        .IgnoreBlank = True ' 空白セルを無視
        .InCellDropdown = True ' ドロップダウンリストを表示
        .InputTitle = "発注日付"
        .ErrorTitle = "入力エラー"
        .InputMessage = "日付を入力してください。"
        .ErrorMessage = "入力値が有効な日付ではありません。"
        .ShowInput = True ' 入力メッセージを表示
        .ShowError = True ' エラーメッセージを表示
    End With
End Sub

'共有フォルダへのアクセス権限があるかチェック
Public Sub CheckDirPermission()
    Dim data As New DataSheetAccesser
    If Not CheckDirectoryAccess(data.SaveDirPath) Then
        MsgBox "共有フォルダへのアクセス権限がありません。使用するには情報課へ依頼してください"
        End
    End If
End Sub


'入力値の動的なバリデーションチェック

'部門コードの入力値をチェック
Public Sub CheckExistsBumon(bumonCode As Variant)
    
    '空の場合
    If IsEmpty(bumonCode) Then
        End
    End If
    
    '数値でない場合
    If Not IsNumeric(bumonCode) Then
        MsgBox ("数値を入力して下さい")
        End
    End If

    '部門コードが存在するか
    Dim dataStorage As New DataBaseAccesser
    If Not dataStorage.ExistsBumon(bumonCode) Then
        MsgBox ("正しい部門コードを入力して下さい")
        End
    End If
    
End Sub

'発注日の入力値をチェック
Public Sub CheckDateFormat(targetDate As Variant)
    
    '空の場合
    If IsEmpty(targetDate) Then
        End
    End If
    
    '日付でない場合
    If Not IsDate(targetDate) Then
        MsgBox ("日付を入力して下さい")
        End
    End If
    
End Sub
