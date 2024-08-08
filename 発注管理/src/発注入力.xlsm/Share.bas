Attribute VB_Name = "Share"

Public Function GetRangeValue(rng As range) As Collection
    Dim cell As range
    Dim col As New Collection
    
    ' �͈͓��̊e�Z�������[�v
    For Each cell In rng
        col.add cell.value
    Next cell
    
    Set GetRangeValue = col
    
End Function

'�������A���t�@�x�b�g�ɕύX
Function NumberToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        NumberToLetter = "Out of Range"
    Else
        NumberToLetter = Chr(64 + num)
    End If
End Function

'�V�[�g�ɗ���w�肵�ē���
Sub writeData(ws As Worksheet, rowIndex As Long, colIndex As Integer, writeData As Collection)
    Dim item As Variant

    For Each item In writeData
        ws.Cells(rowIndex, colIndex).value = item
        rowIndex = rowIndex + 1
    Next item

End Sub

'collection�^�̕ϐ����׏d������l�����O
Function FilterCollection(baseCol As Collection, filterCol As Collection) As Collection
    Dim resultCol As New Collection
    Dim itemBase As Variant
    Dim itemFilter As Variant
    Dim exists As Boolean
    
    ' baseCol�̒l�����[�v���āAfilterCol�ɑ��݂��Ȃ����̂���resultCol�ɒǉ�
    For Each itemBase In baseCol
        exists = False
        For Each itemFilter In filterCol
            If itemBase = itemFilter Then
                exists = True
                Exit For
            End If
        Next itemFilter
        If Not exists Then
            resultCol.add itemBase
        End If
    Next itemBase
    
    ' ���ʂ̃R���N�V������Ԃ�
    Set FilterCollection = resultCol
End Function

'Number�� MultipleOf�̔{���̏ꍇ��True��Ԃ�
Function IsMultiple(Number As Long, MultipleOf As Long) As Boolean
    If MultipleOf = 0 Then
        IsMultiple = False
    Else
        IsMultiple = (Number Mod MultipleOf = 0)
    End If
End Function


