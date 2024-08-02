Attribute VB_Name = "Module1"
Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim dataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New fileFilter
    Dim users As New UserCodeAccesser
    
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
    
    Dim UserCodes As Collection
    Set UserCodes = users.GetEmployeeCodes(load.BumonCode)
    
    
    For Each UserCode In UserCodes
    
        '�]�ƈ����̎擾
        Dim userName As String
        userName = dataStorage.GetUserName(UserCode)
        
        '�t�@�C���̒��o����������̐ݒ�
        UserCodeFilter = fileProperty.UserCodeIdentifier & UserCode & fileProperty.BreakIdentifier
        
        '�t�B���^�[�̎��s
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        '�t�B���^�[�̌��ʂ����݂���ꍇ
        If filePathCollection.Count > 0 Then
            fileProperty.filePath = data.DataDirPath & "\" & filePathCollection(1)
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next UserCode
    
End Sub

