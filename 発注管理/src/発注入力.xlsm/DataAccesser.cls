VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccesser"
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
    
    If Not Result.EOF And Result.Fields(0).Value > 0 Then
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
    query = "SELECT COUNT(1) FROM dbo.M部門 WHERE 部門コード = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門コード", BumonCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    If Not Result.EOF And Result.Fields(0).Value > 0 Then
        ExistsBumon = True
    Else
        ExistsBumon = False
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
        GetUserName = Result.Fields("担当者名").Value
    Else
        GetUserName = ""
    End If
    
    ' Recordsetを閉じる
    Result.Close
    Set Result = Nothing
    Set Executor = Nothing
End Function

' 部門の全ての商品取得
Public Function GetAllProducts(ByVal BumonCD As Integer) As Object
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT * FROM HM商品 WHERE 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "部門CD", BumonCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetAllProducts = Result
End Function

' 部門の条件に合う商品取得
Public Function GetFilterProducts(ByVal BumonCD As Integer, ByVal ColumnName As String, ByVal Value As String) As Object
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT * FROM HM商品 WHERE " & ColumnName & " = ? AND 部門CD = ?"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")
    parameters.Add "条件", Value
    parameters.Add "部門CD", BumonCD

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetFilterProducts = Result
End Function

' 商品マスターの列名取得
Public Function GetTableColumn() As Object
    Dim Executor As SqlExecutor
    Set Executor = New SqlExecutor
    Executor.SetConnection MFRE05ConnectionString
    
    Dim query As String
    query = "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'HM商品'"
    
    Dim parameters As Object
    Set parameters = CreateObject("Scripting.Dictionary")

    Dim Result As Object
    Set Result = Executor.ExecuteSelect(query, parameters)
    
    Set GetTableColumn = Result
End Function

