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
        resultDict(key) = dict1(key)
    Next key

    ' dict2�̓��e��resultDict�ɒǉ�
    For Each key In dict2.Keys
        If resultDict.exists(key) Then
            ' �����̒l�����l�̏ꍇ�͉��Z
            If IsNumeric(resultDict(key)) And IsNumeric(dict2(key)) Then
                resultDict(key) = resultDict(key) + dict2(key)
            ' �Е������l�ŕЕ������l�łȂ��ꍇ�͐��l�̕������ʂɔ��f
            ElseIf IsNumeric(resultDict(key)) Then
                resultDict(key) = resultDict(key)
            ElseIf IsNumeric(dict2(key)) Then
                resultDict(key) = dict2(key)
            ' �����Ƃ����l�łȂ��ꍇ�͋󕶎������ʂɔ��f
            Else
                resultDict(key) = ""
            End If
        Else
            resultDict(key) = dict2(key) ' �L�[�����݂��Ȃ��ꍇ�A�V�����ǉ�
        End If
    Next key

    Set MergeDictionaries = resultDict
End Function

'�񎟌��z�񂩂��s�ڂ��폜
Function RemoveFirstRow(ByVal arr As Variant) As Variant
    Dim newArr() As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long, j As Long
    
    ' �z��̃T�C�Y���擾
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' �V�����z��̃T�C�Y��ݒ�
    ReDim newArr(1 To numRows - 1, 1 To numCols)
    
    ' ��s�ڂ��폜���ĐV�����z��ɃR�s�[
    For i = 2 To numRows
        For j = 1 To numCols
            newArr(i - 1, j) = arr(i, j)
        Next j
    Next i
    
    ' �V�����z���Ԃ�
    RemoveFirstRow = newArr
End Function
