Attribute VB_Name = "Share"

Public Function GetRangeValue(rng As range) As Collection
    Dim cell As range
    Dim col As New Collection
    
    ' 範囲内の各セルをループ
    For Each cell In rng
        col.add cell.value
    Next cell
    
    Set GetRangeValue = col
    
End Function

'数字をアルファベットに変更
Private Function NumberToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        NumberToLetter = "Out of Range"
    Else
        NumberToLetter = Chr(64 + num)
    End If
End Function
