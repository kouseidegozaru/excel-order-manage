Attribute VB_Name = "LoadOrderData"
Sub Loads()
    Application.ScreenUpdating = False
    LoadFileProperty
    LoadData
    Application.ScreenUpdating = True
End Sub

Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim DataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
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
    
    '������w�肵�ď]�ƈ��R�[�h�̎擾
    Dim UserCodes As Collection
    Set UserCodes = users.GetEmployeeCodes(load.BumonCode)
    
    
    For Each UserCode In UserCodes
    
        '�]�ƈ����̎擾
        Dim userName As String
        userName = DataStorage.GetUserName(UserCode)
        
        '�t�@�C���̒��o����������̐ݒ�
        UserCodeFilter = fileProperty.UserCodeIdentifier & UserCode & fileProperty.BreakIdentifier
        
        '�t�B���^�[�̎��s
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter, UserCodeFilter)
        
        '�t�B���^�[�̌��ʂ����݂���ꍇ
        If filePathCollection.Count > 0 Then
            '�t�@�C�����̎擾����
            fileProperty.filePath = data.DataDirPath & "\" & filePathCollection(1)
            '�t�@�C���v���p�e�B�̕\��
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            '�t�@�C���v���p�e�B�̕\��
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next UserCode
    
End Sub

Sub LoadData()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    
    load.ClearData
    
'''�������̎擾'''
    Dim BumonCodeFilter As String
    Dim TargetDateFilter As String
    
    '�t�@�C���̒��o����������̐ݒ�
    BumonCodeFilter = fileProperty.BumonCodeIdentifier & load.BumonCode & fileProperty.BreakIdentifier
    TargetDateFilter = fileProperty.DateIdentifier & Format(load.TargetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '�t�B���^�[�̎��s
    Dim filePathCollection As Collection
    filter.DirPath = data.DataDirPath
    Set filePathCollection = filter.AndFilter(BumonCodeFilter, TargetDateFilter)
    
    '���i�R�[�h���L�[�l�A���ʂ�l�Ƃ��������^�ϐ�
    Dim resultDict As New Scripting.Dictionary
    
    For Each FileName In filePathCollection
        
        '�f�[�^�̎擾����
        data.FileName = FileName
        '�t�@�C�����̎擾����
        fileProperty.filePath = data.filePath
        '���i�������
        load.WriteAllData data.dataNoHeader
        '�f�[�^���[�N�u�b�N�����
        data.CloseWorkBook
        
    Next FileName
    
    Dim rs As ADODB.Recordset
    Set rs = load.ProductsGroupData
    load.ClearData
    load.WriteAllData RemoveFirstRow(RecordsetToArray(rs))
    
End Sub
