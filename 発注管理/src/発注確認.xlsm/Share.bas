Attribute VB_Name = "Share"

'�������A���t�@�x�b�g�ɕύX
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function


'''�ȉ���SheetAccesser�݂̂Ŏg�p���鍀��'''

' range�Ŏw�肵���͈͂���s�܂��͈��̏ꍇ�Ɉꎟ����Collection�Ɋi�[����
Public Function RangeToOneDimCollection(rng As Range) As Collection
    Dim arr As Variant
    Dim oneDimCollection As New Collection
    Dim i As Integer

    arr = rng.value
    
    If IsEmpty(arr) Then
        Set RangeToOneDimCollection = oneDimCollection
        Exit Function
    End If

    ' ��s����񂩂𔻒�
    If rng.Rows.Count = 1 Then
        ' ��s�̏ꍇ
        For i = 1 To rng.Columns.Count
            oneDimCollection.Add arr(1, i)
        Next i
    ElseIf rng.Columns.Count = 1 Then
        ' ���̏ꍇ
        For i = 1 To rng.Rows.Count
            oneDimCollection.Add arr(i, 1)
        Next i
    End If

    Set RangeToOneDimCollection = oneDimCollection
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
    If col.Count = 0 Then
        Set RemoveFirstRow = newCol
        Exit Function
    End If

    ' �ŏ��̍s���폜����
    numRows = col.Count
    
    ' �ŏ��̍s���폜���ĐV����Collection�ɃR�s�[
    For i = 2 To numRows
        Set row = New Collection
        For j = 1 To col(i).Count
            row.Add col(i)(j)
        Next j
        newCol.Add row
    Next i
    
    ' �V����Collection��Ԃ�
    Set RemoveFirstRow = newCol
End Function

'�񎟌��z���񎟌��̃R���N�V�����ɕϊ�����
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

