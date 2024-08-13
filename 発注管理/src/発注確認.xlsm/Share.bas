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


Function RecordsetToCollection(rs As ADODB.Recordset) As Collection
    Dim col As Collection
    Dim rowCol As Collection
    Dim i As Long
    
    ' �R���N�V������������
    Set col = New Collection
    
    ' ���R�[�h�Z�b�g����łȂ����Ƃ��m�F
    If Not rs.EOF Then
        rs.MoveFirst
        
        ' �f�[�^���R���N�V�����Ɋi�[
        Do Until rs.EOF
            Set rowCol = New Collection
            For i = 0 To rs.Fields.Count - 1
                rowCol.Add rs.Fields(i).value, rs.Fields(i).name
            Next i
            col.Add rowCol
            rs.MoveNext
        Loop
    End If
    
    Set RecordsetToCollection = col
End Function


