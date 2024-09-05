Attribute VB_Name = "Share"

'�������A���t�@�x�b�g�ɕύX
Function IndexToLetter(ByVal num As Integer) As String
    If num < 1 Or num > 26 Then
        IndexToLetter = "Out of Range"
    Else
        IndexToLetter = Chr(64 + num)
    End If
End Function

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

'�f�B���N�g���ւ̃A�N�Z�X�����̃`�F�b�N
Function CheckDirectoryAccess(ByVal directoryPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Dir�֐��Ńf�B���N�g�������݂��A�A�N�Z�X�\���ǂ����m�F
    If Dir(directoryPath, vbDirectory) <> "" Then
        CheckDirectoryAccess = True
    Else
        CheckDirectoryAccess = False
    End If
    
    Exit Function

ErrorHandler:
    ' �G���[�n���h�����O: �A�N�Z�X�ł��Ȃ��ꍇ�AFalse��Ԃ�
    CheckDirectoryAccess = False
End Function

'''�ȉ���SheetAccesser�݂̂Ŏg�p���鍀��'''

' range�Ŏw�肵���͈͂���s�܂��͈��̏ꍇ�Ɉꎟ����Collection�Ɋi�[����
Public Function RangeToOneDimCollection(rng As Range) As Collection
    Dim arr As Variant
    Dim oneDimCollection As New Collection
    Dim i As Integer

    arr = rng.value
    
    '��̏ꍇ��̃R���N�V������Ԃ�
    If IsEmpty(arr) Then
        Set RangeToOneDimCollection = oneDimCollection
        Exit Function
    End If
    
    'range.value�Ŕ͈͂���̃Z���Ԓn�݂̂��w���ꍇ�ɔz��ł͂Ȃ��Ȃ�
    If Not IsArray(arr) Then
        oneDimCollection.Add arr
    
    ' ��s����񂩂𔻒�
    ElseIf rng.Rows.count = 1 Then
        ' ��s�̏ꍇ
        For i = 1 To rng.Columns.count
            oneDimCollection.Add arr(1, i)
        Next i
        
    ElseIf rng.Columns.count = 1 Then
        ' ���̏ꍇ
        For i = 1 To rng.Rows.count
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
