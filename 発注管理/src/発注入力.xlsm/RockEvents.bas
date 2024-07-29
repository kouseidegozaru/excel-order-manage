Attribute VB_Name = "RockEvents"
'変更処理が呼ばれるとその変更を検知してまた変更が起きる連鎖が生まれてしまうので
'一つの変更処理が終わるまで検知を無視する

Public isIgnoreChange As Boolean

Sub SetIgnoreState(isIgnore As Boolean)
    isIgnoreChange = isIgnore
    
End Sub

Function GetIgnoreState() As Boolean
    GetIgnoreState = isIgnoreChange
    
End Function
