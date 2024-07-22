VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private connectionString As String

' コンストラクタ
Public Sub Class_Initialize(connString As String)
    
End Sub

Public Function SetConnection(connString As String)
    connectionString = connString
End Function
' SELECT文を実行してDataTableを返す
Public Function ExecuteSelect(Query As String, Optional Parameters As Scripting.Dictionary) As Variant
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim dt As Object
    
    Set conn = New ADODB.Connection
    conn.connectionString = connectionString
    conn.Open
    
    Set rs = New ADODB.Recordset
    With rs
        .ActiveConnection = conn
        .Open Query
        Set dt = CreateObject("System.Data.DataTable")
        dt.Load rs
    End With
    
    conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
    ExecuteSelect = dt
End Function

' INSERT文を実行して影響を受けた行数を返す
Public Function ExecuteInsert(Query As String, Optional Parameters As Scripting.Dictionary) As Long
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim affectedRows As Long
    
    Set conn = New ADODB.Connection
    conn.connectionString = connectionString
    conn.Open
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = Query
    
    If Not Parameters Is Nothing Then
        Dim key As Variant
        For Each key In Parameters.Keys
            cmd.Parameters.Append cmd.CreateParameter(key, adVarChar, adParamInput, , Parameters(key))
        Next
    End If
    
    affectedRows = cmd.Execute
    
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    
    ExecuteInsert = affectedRows
End Function

