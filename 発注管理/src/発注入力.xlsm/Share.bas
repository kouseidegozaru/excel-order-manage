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

' �f�[�^���V�[�g�ɏ�������
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCount As Long, colCount As Long

    ' writeData���z�񂩃R���N�V���������m�F
    If IsArray(writeData) Then
        ' �z��̏ꍇ
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + i, startColIndex + j).value = writeData(i, j)
            Next j
        Next i
    ElseIf TypeName(writeData) = "Collection" Then
        ' �R���N�V�����̏ꍇ
        rowCount = 0
        For Each item In writeData
            If IsArray(item) Then
                ' �����z��̒������擾
                colCount = UBound(item, 2) - LBound(item, 2) + 1
                For j = LBound(item, 2) To UBound(item, 2)
                    ws.Cells(startRowIndex + rowCount, startColIndex + j - LBound(item, 2)).value = item(j)
                Next j
                rowCount = rowCount + 1
            End If
        Next item
    Else
        ' �G���[�n���h�����O
        Err.Raise vbObjectError + 9999, "writeData", "writeData�͔z��܂��̓R���N�V�����łȂ���΂Ȃ�܂���B"
    End If
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
        IsMultiple = True
    Else
        IsMultiple = (Number Mod MultipleOf = 0)
    End If
End Function


