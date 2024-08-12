Attribute VB_Name = "Share"

'�������A���t�@�x�b�g�ɕύX
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function

' �f�[�^���V�[�g�ɏ�������
Sub writeData(ws As Worksheet, startRowIndex As Long, startColIndex As Integer, writeData As Variant)
    Dim i As Long, j As Long
    Dim item As Variant
    Dim rowCollection As Variant
    Dim rowCount As Long, colCount As Long

    ' writeData���z�񂩃R���N�V���������m�F
    If IsArray(writeData) Then
        ' �z��̏ꍇ
        For i = LBound(writeData, 1) To UBound(writeData, 1)
            For j = LBound(writeData, 2) To UBound(writeData, 2)
                ws.Cells(startRowIndex + i - 1, startColIndex + j - 1).value = writeData(i, j)
            Next j
        Next i
    ElseIf TypeName(writeData) = "Collection" Then
        ' �R���N�V�����̏ꍇ
        rowCount = 0
        For Each rowCollection In writeData
            colCount = 0
            If TypeName(rowCollection) = "Collection" Then
                ' ����������ɃR���N�V�����̏ꍇ
                For Each item In rowCollection
                    ws.Cells(startRowIndex + rowCount, startColIndex + colCount).value = item
                    colCount = colCount + 1
                Next item
            ElseIf IsArray(rowCollection) Then
                ' �������z��̏ꍇ
                For j = LBound(rowCollection, 1) To UBound(rowCollection, 1)
                    ws.Cells(startRowIndex + rowCount, startColIndex + j - LBound(rowCollection, 1)).value = rowCollection(j)
                Next j
            End If
            rowCount = rowCount + 1
        Next rowCollection
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


'Number�� MultipleOf�̔{���̏ꍇ��True��Ԃ�
Function IsMultiple(Number As Long, MultipleOf As Long) As Boolean
    If MultipleOf = 0 Then
        IsMultiple = True
    Else
        IsMultiple = (Number Mod MultipleOf = 0)
    End If
End Function

'�񎟌��R���N�V���������s�ڂ��폜
Function RemoveFirstRow(ByVal col As Collection) As Collection
    Dim newCol As Collection
    Dim item As Variant
    Dim row As Collection
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long, j As Long
    
    ' �V����Collection���쐬
    Set newCol = New Collection
    
    ' Collection�̍ŏ��̍s�����o��
    If col.count = 0 Then
        Set RemoveFirstRow = newCol
        Exit Function
    End If

    ' �ŏ��̍s���폜����
    numRows = col.count
    
    ' �ŏ��̍s���폜���ĐV����Collection�ɃR�s�[
    For i = 2 To numRows
        Set row = New Collection
        For j = 1 To col(i).count
            row.Add col(i)(j)
        Next j
        newCol.Add row
    Next i
    
    ' �V����Collection��Ԃ�
    Set RemoveFirstRow = newCol
End Function


Function ArrayToCollection(ByVal arr As Variant) As Collection
    Dim col As New Collection
    Dim innerCol As Collection
    Dim i As Long, j As Long

    ' �s�̃��[�v
    For i = LBound(arr, 1) To UBound(arr, 1)
        Set innerCol = New Collection
        
        ' ��̃��[�v
        For j = LBound(arr, 2) To UBound(arr, 2)
            innerCol.Add arr(i, j)
        Next j
        
        col.Add innerCol
    Next i
    
    Set ArrayToCollection = col
End Function

