Attribute VB_Name = "Module1"
Sub ボタン2_Click()
    Dim dataAccesser As New dataAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = 40 ' 例として1を使用
    
    ' データベースからレコードセットを取得
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    Dim exporter As New ProductsSearchSheet
    exporter.Initialize rs
    exporter.ExportRecordSet
    
    ' チェックされたIDを取得
    Dim checkedIDs As Collection
    Set checkedIDs = exporter.GetCheckedValue(1)
    
    Dim id As Variant
    For Each id In checkedIDs
        Debug.Print "Checked ID: " & id
    Next id
End Sub
