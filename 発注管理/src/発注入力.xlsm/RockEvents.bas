Attribute VB_Name = "RockEvents"
'変更処理が呼ばれるとその変更を検知してまた変更が起きる連鎖が生まれてしまうので
'一つの変更処理が終わるまで検知を無視する

'このモジュールではステータスのみを管理
Public isIgnoreChange As Boolean

Public Property Let IsIgnoreChangeEvents(isIgnore As Boolean)
    isIgnoreChange = isIgnore
End Property

Public Property Get IsIgnoreChangeEvents() As Boolean
    IsIgnoreChangeEvents = isIgnoreChange
End Property
