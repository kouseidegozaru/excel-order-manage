VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProductsSearchForm 
   Caption         =   "���i�����t�H�[��"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   OleObjectBlob   =   "ProductsSearchForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
    BumonCD = 40 ' ��Ƃ���1���g�p
    
    ' �f�[�^�x�[�X���烌�R�[�h�Z�b�g���擾
    Set rs = dataAccesser.GetAllProducts(BumonCD)
    
    ' ���R�[�h�Z�b�g��SetProductsTable�ɓn���ĕ\��
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

    '��
    Dim ColumnCount As Integer
    ColumnCount = ProductsDataTable.Fields.Count
    
    '�񖼂��擾���Ĕz��Ɋi�[
    Dim ColumnNames() As String
    ReDim ColumnNames(ColumnCount - 1)
    Dim i As Integer
    For i = 0 To ColumnCount - 1
        ColumnNames(i) = ProductsDataTable.Fields(i).Name
    Next i
    
    '�s��
    Dim RowCnt As Integer
    RowCnt = 0
    
    With Me.ProductsTable
        ProductsDataTable.MoveFirst '�ŏ��̃��R�[�h�Ɉړ�
        .ColumnCount = ColumnCount

        '�񖼂��ŏ��ɒǉ�
        .AddItem Join(ColumnNames, vbTab)
        
        '�f�[�^��ǉ�
        .AddItem "" '�󗓂̃��X�g��ǉ�
        .Column = ProductsDataTable.GetRows
    End With
    
    ' ���R�[�h�Z�b�g�����
    ProductsDataTable.Close
    Set ProductsDataTable = Nothing
End Sub

