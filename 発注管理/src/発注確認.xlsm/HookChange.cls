VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HookChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'セルの変更を検知するための汎用クラス

Private targetRange As Range
Private isIgnore As Boolean

Public Sub Class_Initialize()
    isIgnore = False
End Sub

'変更されたシートの範囲をセット
Public Sub SetRange(targetRng As Range)
    Set targetRange = targetRng
End Sub

'シートの特定の範囲の変更があったか
Public Function IsChanged(rng As Range) As Boolean
    IsChanged = False
    
    '監視しない
    If isIgnore Then
        Exit Function
    End If
    
    If Not Intersect(targetRange, rng) Is Nothing Then
        IsChanged = True
    End If

End Function

'行が削除されたか
Public Function IsDeleted(lastRowCnt As Long, currentRowCnt As Long) As Boolean
    IsDeleted = False

    '監視しない
    If isIgnore Then
        Exit Function
    End If
    
    '行数が減少した場合、削除されたと判断する
    If lastRowCnt > currentRowCnt Then
        IsDeleted = True
    End If

End Function

'シートの変更があっても無視する
Public Sub Ignore(isChangeIgnore As Boolean)
    isIgnore = isChangeIgnore
End Sub


