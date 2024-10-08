Attribute VB_Name = "Validations"
Sub SetValidations()
    'セルにバリデーションチェックを設定
    SetBumonCD
    SetUserCD
    SetDate
End Sub

Private Sub SetBumonCD()
    Dim order As New OrderSheetAccesser
    Dim rng As Range
    Set rng = order.BumonCodeRange
    
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
    Dim order As New OrderSheetAccesser
    Dim rng As Range
    Set rng = order.UserCodeRange
    
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
    Dim order As New OrderSheetAccesser
    Dim rng As Range
    Set rng = order.TargetDateRange
    
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

'合わせ数のバリデーションチェック
Public Function IsMatchQty() As Boolean
    Dim order As New OrderSheetAccesser
    Dim i As Long
    
    '数量
    Dim qtyCol As Collection
    '合わせ数
    Dim matchCol As Collection
    
    Set qtyCol = order.qty
    Set matchCol = order.match
    
    IsMatchQty = True
    
    For i = 1 To matchCol.count
        If Not IsMultiple(qtyCol(i), matchCol(i)) Then
            IsMatchQty = False
            Exit For
        End If
    Next i
    
End Function

'共有フォルダへのアクセス権限があるかチェック
Public Sub CheckDirPermission()
    Dim data As New DataSheetAccesser
    If Not CheckDirectoryAccess(data.SaveDirPath) Then
        MsgBox "共有フォルダへのアクセス権限がありません。使用するには情報課へ依頼してください"
        End
    End If
End Sub


'入力値の動的なバリデーションチェック

'担当者コードの入力値をチェック
Public Sub CheckExistsUser(userCode As Variant)
    
    '空の場合
    If IsEmpty(userCode) Then
        End
    End If
    
    '数値でない場合
    If Not IsNumeric(userCode) Then
        MsgBox ("数値を入力して下さい")
        End
    End If

    '担当者コードが存在するか
    Dim dataStorage As New DataBaseAccesser
    If Not dataStorage.ExistsUser(userCode) Then
        MsgBox ("正しい担当者コードを入力して下さい")
        End
    End If
    
End Sub

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
    
    '現在の時刻が9時を過ぎている場合
    If Time >= #9:00:00 AM# Then
        '発注日が今日の日付の場合
        If DateValue(targetDate) = Date Then
            MsgBox ("現在の時刻が9時を過ぎているため、今日の日付の発注はできません。")
            End
        End If
    End If
    
End Sub
