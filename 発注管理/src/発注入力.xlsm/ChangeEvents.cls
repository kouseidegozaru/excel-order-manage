VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangeEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'セルの変更を検知するための汎用クラス

Private ws As Worksheet
Private targetRange As range
Private isIgnore As Boolean
Private lastUpdatedRowCnt As Long

Public Sub Initialize(targetWs As Worksheet)
    If targetWs Is Nothing Then
        MsgBox "Invalid worksheet provided.", vbCritical
        Exit Sub
    End If
    Set ws = targetWs
    isIgnore = False
End Sub

Public Sub SetRange(targetRng As range)
    If targetRng Is Nothing Then
        MsgBox "Invalid range provided.", vbCritical
        Exit Sub
    End If
    Set targetRange = targetRng
End Sub

Public Function IsChanged(strRange As String) As Boolean
    IsChanged = False
    
    '監視しない
    If isIgnore Then
        Exit Function
    End If
    
    If ws Is Nothing Then
        MsgBox "Worksheet is not initialized.", vbCritical
        Exit Function
    End If
    
    If Not Intersect(targetRange, ws.range(strRange)) Is Nothing Then
        IsChanged = True
    End If

End Function

Public Function IsDeleted(lastRowCnt As Long, currentRowCnt As Long) As Boolean
    IsDeleted = False

    '監視しない
    If isIgnore Then
        Exit Function
    End If
    
    If ws Is Nothing Then
        MsgBox "Worksheet is not initialized.", vbCritical
        Exit Function
    End If
    
    '行数が減少した場合、削除されたと判断する
    If lastRowCnt > currentRowCnt Then
        IsDeleted = True
    End If

End Function

Public Sub Ignore(isChangeIgnore As Boolean)
    isIgnore = isChangeIgnore
End Sub

