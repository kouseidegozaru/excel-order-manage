VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProductsSearchForm 
   Caption         =   "商品検索フォーム"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   OleObjectBlob   =   "ProductsSearchForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ProductsSearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txbValue_Change()

End Sub

Private Sub UserForm_Activate()
    Dim dataAccesser As New dataAccesser
    Dim rs As ADODB.recordSet
    Dim BumonCD As Integer
    BumonCD = 40 ' 例として1を使用
    
    ' データベースからレコードセットを取得
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    ' レコードセットをSetProductsTableに渡して表示
    Call SetProductsTable(rs)
End Sub

Private Function NullCheck(Value As Variant) As String
    If IsNull(Value) Then
         NullCheck = ""
    Else
         NullCheck = Value
    End If
End Function

Private Sub SetProductsTable(ProductsDataTable As ADODB.recordSet)

    '列数
    Dim ColumnCount As Integer
    ColumnCount = ProductsDataTable.Fields.Count
    
    '列名を取得して配列に格納
    Dim ColumnNames() As String
    ReDim ColumnNames(ColumnCount - 1)
    Dim i As Integer
    For i = 0 To ColumnCount - 1
        ColumnNames(i) = ProductsDataTable.Fields(i).Name
    Next i
    
    '行数
    Dim RowCnt As Integer
    RowCnt = 0
    
    With Me.ProductsTable
        ProductsDataTable.MoveFirst '最初のレコードに移動
        .ColumnCount = ColumnCount

        '列名を最初に追加
        .AddItem Join(ColumnNames, vbTab)
        
        'データを追加
        .AddItem "" '空欄のリストを追加
        .Column = ProductsDataTable.GetRows
    End With
    
    ' レコードセットを閉じる
    ProductsDataTable.Close
    Set ProductsDataTable = Nothing
End Sub

