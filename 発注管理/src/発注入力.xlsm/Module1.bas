Attribute VB_Name = "Module1"
Sub �{�^��2_Click()
    Dim dataAccesser As New dataAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = 40 ' ��Ƃ���1���g�p
    
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    Dim exporter As New ProductsSearchSheet
    exporter.Initialize rs
    exporter.ExportRecordSet
    
    ' �`�F�b�N���ꂽID���擾
    Dim checkedIDs As Collection
    Set checkedIDs = exporter.GetCheckedValue(1)
    
    Dim id As Variant
    For Each id In checkedIDs
        Debug.Print "Checked ID: " & id
    Next id
End Sub
