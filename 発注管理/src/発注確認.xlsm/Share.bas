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

' �f�[�^���V�[�g�ɏ�������
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCount As Long, colCount As Long

    ' writeData���z�񂩃R���N�V���������m�F
    If IsArray(writeData) Then
        rowcnt = 0
        ' �z��̏ꍇ
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            colcnt = 0
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + rowcnt, startColIndex + colcnt).value = writeData(i, j)
                colcnt = colcnt + 1
            Next j
            rowcnt = rowcnt + 1
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
    Dim minRows As Long
    Dim minCols As Long
    Dim i As Long, j As Long
    
    ' �z��̃T�C�Y���擾
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    minRows = LBound(arr, 1)
    minCols = LBound(arr, 2)
    
    ' �V�����z��̃T�C�Y��ݒ�
    ReDim newArr(minRows To numRows - 1, minCols To numCols)
    
    ' ��s�ڂ��폜���ĐV�����z��ɃR�s�[
    For i = minRows + 1 To numRows
        For j = minCols To numCols
            newArr(i - 1, j) = arr(i, j)
        Next j
    Next i
    
    ' �V�����z���Ԃ�
    RemoveFirstRow = newArr
End Function

Function RecordsetToArray(rs As ADODB.Recordset) As Variant
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim colCount As Long
    
    ' ���R�[�h�Z�b�g�̗񐔂��擾
    colCount = rs.Fields.Count
    
    ' ���R�[�h�Z�b�g�̍s�����擾
    rs.MoveLast
    rowCount = rs.RecordCount
    rs.MoveFirst
    
    ' �񎟌��z���������
    ReDim arr(0 To rowCount, 0 To colCount - 1)
    
    ' �w�b�_�[��z��Ɋi�[
    For i = 0 To colCount - 1
        arr(0, i) = rs.Fields(i).name
    Next i
    
    ' �f�[�^��z��Ɋi�[
    i = 1
    Do Until rs.EOF
        For j = 0 To colCount - 1
            arr(i, j) = rs.Fields(j).value
        Next j
        rs.MoveNext
        i = i + 1
    Loop
    
    RecordsetToArray = arr
End Function

