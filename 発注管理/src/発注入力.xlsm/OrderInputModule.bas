Attribute VB_Name = "OrderInputModule"
'�����t�H�[���Ăяo��
Sub Search()
    Dim dataAccesser As New dataAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    Dim exporter As New ProductsSearchSheet
    exporter.Initialize rs
    exporter.ExportRecordSet
    
End Sub

