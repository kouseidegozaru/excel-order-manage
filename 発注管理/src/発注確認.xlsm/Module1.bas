Attribute VB_Name = "Module1"
Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim dataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New fileFilter
    
    Dim BumonCodeFilter As String
    Dim UserCodeFilter As String
    Dim TargetDateFilter As String
    
    '�v���p�e�B���N���A
    load.ClearFileProperty
    
    '�t�@�C���̒��o����������̐ݒ�
    BumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    TargetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '�T������f�B���N�g���̐ݒ�
    filter.DirPath = data.DataDirPath
    
    Dim rs As ADODB.Recordset
    Set rs = dataStorage.GetUserCodes(load.BumonCode)
    
    Do Until rs.EOF
        '�t�@�C���̒��o����������̐ݒ�
        UserCodeFilter = fileProperty.UserCodeIdentifier & rs.Fields("�S����CD").value & fileProperty.BreakIdentifier
        
        '�t�B���^�[�̎��s
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        '�t�B���^�[�̌��ʂ����݂���ꍇ
        If filePathCollection.Count > 0 Then
            fileProperty.filePath = data.DataDirPath & "\" & filePathCollection(1)
            load.AddFileProperty rs.Fields("�S���Җ�").value, True, fileProperty.UpdatedDate
        Else
            load.AddFileProperty rs.Fields("�S���Җ�").value, False, Date
        End If
        
        rs.MoveNext
    Loop
    
    
    
    
End Sub

Sub LoadFilePropertys()
Dim filter As New fileFilter
filter.DirPath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\�����Ǘ�\data"

Dim c As Collection
Set c = filter.AndFilter("b40-", "d20240725-", "u30-")
End Sub

