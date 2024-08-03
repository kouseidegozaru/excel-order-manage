Attribute VB_Name = "Share"

Public Function GetRangeValue(rng As Range) As Collection
    Dim cell As Range
    Dim col As New Collection
    
    ' �͈͓��̊e�Z�������[�v
    For Each cell In rng
        col.Add cell.value
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
            resultCol.Add itemBase
        End If
    Next itemBase
    
    ' ���ʂ̃R���N�V������Ԃ�
    Set FilterCollection = resultCol
End Function

'�������r���L�[���������͉̂��Z
Function MergeDictionaries(dict1 As Scripting.Dictionary, dict2 As Scripting.Dictionary) As Scripting.Dictionary
    Dim resultDict As New Scripting.Dictionary
    Dim key As Variant

    ' dict1�̓��e��resultDict�ɃR�s�[
    For Each key In dict1.Keys
        If IsNumeric(dict1(key)) Then
            resultDict(key) = dict1(key)
        Else
            resultDict(key) = "" ' �����ȕ����̏ꍇ�͋󕶎�
        End If
    Next key

    ' dict2�̓��e��resultDict�ɒǉ�
    For Each key In dict2.Keys
        If resultDict.exists(key) Then
            If IsNumeric(dict2(key)) Then
                resultDict(key) = dict2(key)
            ElseIf IsNumeric(resultDict(key)) Then
                ' resultDict�̒l�����l�̏ꍇ�A�ύX���Ȃ�
            Else
                resultDict(key) = "" ' ���������ȕ����̏ꍇ�͋󕶎�
            End If
        Else
            If IsNumeric(dict2(key)) Then
                resultDict(key) = dict2(key)
            Else
                resultDict(key) = "" ' �����ȕ����̏ꍇ�͋󕶎�
            End If
        End If
    Next key

    Set MergeDictionaries = resultDict
End Function



