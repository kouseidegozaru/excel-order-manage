Attribute VB_Name = "OrderInputModule"
'�����t�H�[���Ăяo��
Sub Search()
    Dim DataBaseAccesser As New DataBaseAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = DataBaseAccesser.GetAllProducts(BumonCD)
    
    Dim exporter As New ProductsSearchSheet
    exporter.Initialize rs
    exporter.ExportRecordSet
    
End Sub

