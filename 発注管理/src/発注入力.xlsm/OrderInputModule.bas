Attribute VB_Name = "OrderInputModule"
'検索フォーム呼び出し
Sub Search()
    Dim dataAccesser As New dataAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = GetBumonCD
    
    ' データベースからレコードセットを取得
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    Dim exporter As New ProductsSearchSheet
    exporter.Initialize rs
    exporter.ExportRecordSet
    
End Sub

