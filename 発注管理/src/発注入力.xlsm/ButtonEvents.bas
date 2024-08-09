Attribute VB_Name = "ButtonEvents"

'���i�����V�[�g�Ɋւ��鏈��
'�m��{�^��(���i�R�[�h�̔��f)
Sub Decide()

    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    Set Search.Workbook = ActiveWorkbook
    Search.InitWorkSheet
    
    '�ύX�̃C�x���g�𖳎�
    IsIgnoreChangeEvents = True
    
    '�d�����鏤�i�R�[�h��r��
    Dim WriteData As Collection
    Set WriteData = FilterCollection(Search.GetCheckedProductsCode, _
                                     order.ProductsCode)
                                     
    Dim startRowIndex As Long
    Dim lastRowIndex As Long
    
    startRowIndex = order.DataNextRowNumber
    lastRowIndex = startRowIndex
    
    '�������͂ɏ��i�R�[�h����
    For i = 1 To WriteData.Count
        order.Cells(lastRowIndex, order.ProductCodeColumnNumber) = WriteData(i)
        lastRowIndex = lastRowIndex + 1
    Next i
    
    '�������͂ɏ��i�R�[�h����͂����͈�
    Dim target As range
    Set target = order.Worksheet.range(order.ProductCodeColumn & startRowIndex & ":" & order.ProductCodeColumn & lastRowIndex)
    
    '���i���\��
    DisplayProductsInfo target
    
    '�ۑ�
    SaveData
    
    IsIgnoreChangeEvents = False
    
    order.Worksheet.Activate
    
End Sub

'�������̓V�[�g�Ɋւ��鏈��
'�����t�H�[���X�V
Sub Update()

    Application.ScreenUpdating = False
    
    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    
    Set Search.Workbook = order.Workbook
    Search.InitWorkSheet
    Search.Clear
    
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = DataBaseAccesser.GetAllProducts(order.BumonCode)
    
    Dim rowIndex As Long
    Dim columnIndex As Integer
    
    rowIndex = Search.DataStartRowNumber
    columnIndex = Search.DataStartColumnNumber
    
    ' �f�[�^�̏�������
    rs.MoveFirst
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.Count - 1
            Search.Cells(rowIndex, i + columnIndex) = rs.Fields(i).value
        Next i
        
        ' �`�F�b�N�{�b�N�X�̒ǉ�
        Search.AddCheckBox rowIndex
        
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop
    
    Search.Worksheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'�`�F�b�N�{�b�N�X�̃N���A
Sub ClearCheckBoxes()
    Dim order As New OrderSheetAccesser
    Dim Search As New SearchSheetAccesser
    
    Set Search.Workbook = order.Workbook
    Search.InitWorkSheet
    Search.ClearCheckBoxes
End Sub

'���i�����V�[�g�̕\��
Sub Search()
    Dim order As New OrderSheetAccesser
    order.FormatSheet.Activate
End Sub

'���M
Sub Post()

    Dim result As VbMsgBoxResult
    '���킹���Ɛ��ʂ̃`�F�b�N
    If Not IsMatchQty Then
        result = MsgBox("���킹���ƈ�v���Ȃ����ʂ�����܂��B���M���܂���?", vbYesNo + vbQuestion, "�m�F")
        If result = vbNo Then
            End
        End If
    End If
        
    SaveData
    MsgBox "�f�[�^�𑗐M���܂���"
End Sub
