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

MFRE05ConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFRE05;UID=sa;PWD=mefure;"
MFRE06ConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFRE06;UID=sa;PWD=mefure;"
MFREDWHConnectionString = "PROVIDER=SQLOLEDB;Data Source=MFH006-MFRE;Initial Catalog=MFREDWH;UID=sa;PWD=mefure;"

'担当者が存在するか
Public Function ExistsUser(ByVal UserID As Integer) As Boolean

    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT COUNT(1) FROM dbo.M担当者 WHERE 担当者CD = @担当者CD"
    
    Dim Parameters As New Scripting.Dictionary
    Parameters.Add "@担当者CD", UserID

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    If Result.Rows.Count > 0 And Convert.ToInt32(Result.Rows(0)(0)) > 0 Then
        ExistsUser = True
    Else
        ExistsUser = False
    End If
    
End Function

'部門が存在するか
Public Function ExistsBumon(ByVal BumonCD As Integer) As Boolean

    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT COUNT(1) FROM dbo.M部門 WHERE 部門コード = @部門コード"
    
    Dim Parameters As New Scripting.Dictionary
    Parameters.Add "@部門コード", BumonCD

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    If Result.Rows.Count > 0 And Convert.ToInt32(Result.Rows(0)(0)) > 0 Then
        ExistsBumon = True
    Else
        ExistsBumon = False
    End If
    
End Function

'担当者名取得
Public Function GetUserName(ByVal UserID As Integer) As String

    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT 担当者名 FROM dbo.M担当者 WHERE 担当者CD = @担当者CD"
    
    Dim Parameters As New Scripting.Dictionary
    Parameters.Add "@担当者CD", UserID

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    GetUser = Result(0)("担当者CD")
    
End Function


'部門の全ての商品取得
Public Function GetAllProducts(BumonCD As Integer) As DataTable
            
    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT * FROM HM商品 WHERE 部門CD = @部門CD"
    
    Dim Parameters As New Scripting.Dictionary
    Parameters.Add "@部門CD", BumonCD

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    GetUser = Result
    
End Function

'部門の条件に合う商品取得
Public Function GetAllProducts(BumonCD As Integer, ColumnName As String, Value As String) As DataTable
            
    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT * FROM HM商品 WHERE @列名 IN(@条件) AND 部門CD = @部門CD"
    
    Dim Parameters As New Scripting.Dictionary
    Parameters.Add "@部門CD", BumonCD
    Parameters.Add "@列名", ColumnName
    Parameters.Add "@条件", Value

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    GetUser = Result
    
End Function

'商品マスターの列名取得
Public Function GetTableColumn() As DataTable
            
    Dim Executor As New SqlExecutor
    Executor.SetConnection (MFRE05ConnectionString)
    
    Dim Query As String
    Query = "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = HM商品"
    
    Dim Parameters As New Scripting.Dictionary

    Dim Result As DataTable
    Result = Executor.ExecuteSelect(Query, Parameters)
    
    GetUser = Result
    
End Function



