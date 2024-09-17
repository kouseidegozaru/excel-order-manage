VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBaseAccesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MFRE05ConnectionString As String
Private MFRE06ConnectionString As String
Private MFREDWHConnectionString As String

' コンストラクタ代わりの初期化メソッド
Public Sub Class_Initialize()
    MFRE05ConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFRE05;UID=sa;PWD=mefure;"
    MFRE06ConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFRE06;UID=sa;PWD=mefure;"
    MFREDWHConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFREDWH;UID=sa;PWD=mefure;"
End Sub

' 担当者が存在するか
Public Function ExistsUser(ByVal UserID As Integer) As Boolean
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) FROM dbo.M担当者 WHERE 担当者CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "担当者CD", UserID

    Dim result As Object
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF And result.Fields(0).value > 0 Then
        ExistsUser = True
    Else
        ExistsUser = False
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function

' 部門が存在するか
Public Function ExistsBumon(ByVal bumonCD As Integer) As Boolean
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) FROM dbo.M部門 WHERE 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD

    Dim result As Object
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF And result.Fields(0).value > 0 Then
        ExistsBumon = True
    Else
        ExistsBumon = False
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function

' 商品が存在するか
Public Function ExistsProducts(ByVal bumonCD As Integer, ByVal ProductsCD As Long) As Boolean
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) FROM dbo.HM商品 WHERE dbo.HM商品.部門CD = ? AND 品CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD
    parameters.Add "品CD", ProductsCD

    Dim result As Object
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF And result.Fields(0).value > 0 Then
        ExistsProducts = True
    Else
        ExistsProducts = False
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function

' 担当者名取得
Public Function GetUserName(ByVal UserID As Integer) As String
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT 担当者名 FROM dbo.M担当者 WHERE 担当者CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "担当者CD", UserID

    Dim result As Object
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF Then
        GetUserName = result.Fields("担当者名").value
    Else
        GetUserName = ""
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function

' 部門名取得
Public Function GetBumonName(ByVal bumonCD As Integer) As String
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT 部門名 FROM dbo.M部門 WHERE 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD

    Dim result As Object
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF Then
        GetBumonName = result.Fields("部門名").value
    Else
        GetBumonName = ""
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function

' 部門の全ての商品取得
Public Function GetAllProducts(ByVal bumonCD As Integer) As ADODB.Recordset
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT " & _
       "dbo.HM商品.入荷時間帯, " & _
       "dbo.HM商品.品CD, " & _
       "dbo.M仕入先.仕入先名 AS 仕入先名, " & _
       "M発注先.仕入先CD AS 発注先コード名, " & _
       "M発注先.仕入先名 AS 発注先名, " & _
       "Mメーカー.仕入先名 AS メーカー名, " & _
       "dbo.HM商品.品名, " & _
       "dbo.HM商品.規格, " & _
       "dbo.HM商品.計区, " & _
       "dbo.HM商品.入数, " & _
       "dbo.HM商品.合わせ数, " & _
       "0 AS 数量, " & _
       "dbo.HM商品.[ケス単位] AS 荷姿, " & _
       "dbo.HM商品.仕入単価, " & _
       "'' AS 仕入金額, " & _
       "dbo.HM商品.LT, " & _
       "dbo.HM商品.商品備考, " & _
       "Mメーカー.仕入先ヨミガナ AS メーカー読み " & _
       "FROM dbo.HM商品 " & _
       "INNER JOIN dbo.M仕入先 ON dbo.HM商品.仕入先CD = dbo.M仕入先.仕入先CD " & _
       "INNER JOIN dbo.M仕入先 AS Mメーカー ON dbo.HM商品.[メーカーCD] = Mメーカー.仕入先CD " & _
       "INNER JOIN dbo.M仕入先 AS M発注先 ON dbo.HM商品.[発注先CD] = M発注先.仕入先CD " & _
       "WHERE dbo.HM商品.部門CD = ? AND dbo.HM商品.廃盤日付 IS NULL " & _
       "ORDER BY Mメーカー.仕入先名 ASC, dbo.HM商品.品名 ASC;"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD

    Dim result As ADODB.Recordset
    Set result = Executor.ExecuteSelect(query, parameters)
    
    Set GetAllProducts = result
End Function

' 部門の商品の件数を取得
Public Function GetAllProductsCount(ByVal bumonCD As Integer) As Long
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) " & _
            "FROM dbo.HM商品 " & _
            "WHERE 部門CD = ? AND 廃盤日付 IS NULL;"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD

    Dim result As ADODB.Recordset
    Set result = Executor.ExecuteSelect(query, parameters)
    
    If Not result.EOF Then
        GetAllProductsCount = result.Fields(0).value
    Else
        GetAllProductsCount = 0
    End If
    
    ' Recordsetを閉じる
    result.Close
    Set result = Nothing
    Set Executor = Nothing
End Function
' 部門の条件に合う商品取得
Public Function GetProduct(ByVal bumonCD As Integer, ProductsCD As Long) As ADODB.Recordset
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT " & _
           "dbo.HM商品.品CD, " & _
           "dbo.HM商品.入荷時間帯, " & _
           "dbo.M仕入先.仕入先名 AS 仕入先名, " & _
           "M発注先.仕入先名 AS 発注先名, " & _
           "Mメーカー.仕入先名 AS メーカー名, " & _
           "dbo.HM商品.品名, " & _
           "dbo.HM商品.規格, " & _
           "dbo.HM商品.計区, " & _
           "dbo.HM商品.入数, " & _
           "dbo.HM商品.合わせ数, " & _
           "0 AS 数量, " & _
           "dbo.HM商品.[ケス単位] AS 荷姿, " & _
           "dbo.HM商品.仕入単価, " & _
           "'' AS 仕入金額, " & _
           "dbo.HM商品.LT, " & _
           "dbo.HM商品.商品備考, " & _
           "Mメーカー.仕入先ヨミガナ AS メーカー読み " & _
           "FROM dbo.HM商品 " & _
           "INNER JOIN dbo.M仕入先 ON dbo.HM商品.仕入先CD = dbo.M仕入先.仕入先CD " & _
           "INNER JOIN dbo.M仕入先 AS Mメーカー ON dbo.HM商品.[メーカーCD] = Mメーカー.仕入先CD " & _
           "INNER JOIN dbo.M仕入先 AS M発注先 ON dbo.HM商品.[発注先CD] = M発注先.仕入先CD " & _
           "WHERE dbo.HM商品.部門CD = ? " & _
           "AND dbo.HM商品.品CD = ? " & _
           "AND dbo.HM商品.廃盤日付 IS NULL;"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", bumonCD
    parameters.Add "品CD", ProductsCD

    Dim result As ADODB.Recordset
    Set result = Executor.ExecuteSelect(query, parameters)
    
    Set GetProduct = result
End Function

' 商品マスターの列名取得
Public Function GetTableColumn() As ADODB.Recordset
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'HM商品'"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")

    Dim result As ADODB.Recordset
    Set result = Executor.ExecuteSelect(query, parameters)
    
    Set GetTableColumn = result
End Function

