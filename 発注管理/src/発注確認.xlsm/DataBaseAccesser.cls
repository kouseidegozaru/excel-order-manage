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

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF And Result.Fields(0).value > 0 Then
        ExistsUser = True
    Else
        ExistsUser = False
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
    Set Executor = Nothing
End Function

' 部門が存在するか
Public Function ExistsBumon(ByVal BumonCD As Integer) As Boolean
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) FROM dbo.M部門 WHERE 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", BumonCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF And Result.Fields(0).value > 0 Then
        ExistsBumon = True
    Else
        ExistsBumon = False
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
    Set Executor = Nothing
End Function

' 商品が存在するか
Public Function ExistsProducts(ByVal BumonCD As Integer, ByVal ProductsCD As Long) As Boolean
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COUNT(1) FROM dbo.HM商品 WHERE dbo.HM商品.部門CD = ? AND 品CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", BumonCD
    parameters.Add "品CD", ProductsCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF And Result.Fields(0).value > 0 Then
        ExistsProducts = True
    Else
        ExistsProducts = False
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
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

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF Then
        GetUserName = Result.Fields("担当者名").value
    Else
        GetUserName = ""
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
    Set Executor = Nothing
End Function

' 部門名取得
Public Function GetBumonName(ByVal BumonCD As Integer) As String
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT 部門名 FROM dbo.M部門 WHERE 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", BumonCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF Then
        GetBumonName = Result.Fields("部門名").value
    Else
        GetBumonName = ""
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
    Set Executor = Nothing
End Function

' 部門の条件に合う商品取得
Public Function GetProduct(ByVal BumonCD As Integer, ProductsCD As Long) As ADODB.Recordset
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT " & _
           "dbo.HM商品.入荷時間帯, " & _
           "dbo.M仕入先.仕入先名 AS 仕入先名, " & _
           "M発注先.仕入先名 AS 発注先名, " & _
           "Mメーカー.仕入先名 AS メーカー名, " & _
           "dbo.HM商品.品名, " & _
           "dbo.HM商品.規格, " & _
           "dbo.HM商品.計区, " & _
           "dbo.HM商品.入数, " & _
           "dbo.HM商品.合わせ数, " & _
           "'' AS 数量, " & _
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
           "WHERE " & _
           "dbo.HM商品.部門CD = ? " & _
           "AND dbo.HM商品.品CD = ? " & _
           "AND dbo.HM商品.廃盤日付 IS NULL;"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", BumonCD
    parameters.Add "品CD", ProductsCD

    Dim Result As ADODB.Recordset
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetProduct = Result
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

    Dim Result As ADODB.Recordset
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetTableColumn = Result
End Function

' 部門の担当者コード取得
Public Function GetUserCodes(bumonCode As Integer) As ADODB.Recordset
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT 担当者CD,担当者名 FROM dbo.M担当者 WHERE 部署CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部署コード", bumonCode

    Dim Result As ADODB.Recordset
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetUserCodes = Result
End Function


