Attribute VB_Name = "LoadOrderData"
Sub Loads()
    Application.ScreenUpdating = False
    LoadFileProperty
    LoadData
    LoadOrderedData
    Application.ScreenUpdating = True
End Sub

Sub LoadFileProperty()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim DataStorage As New DataBaseAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    Dim users As New UserCodeAccesser
    
    Dim bumonCodeFilter As String
    Dim userCodeFilter As String
    Dim targetDateFilter As String
    
    '�v���p�e�B���N���A
    load.ClearFileProperty
    
    '�t�@�C���̒��o����������̐ݒ�
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.bumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.targetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '�T������f�B���N�g���̐ݒ�
    filter.DirPath = data.SaveDirPath
    
    '������w�肵�ď]�ƈ��R�[�h�̎擾
    Dim userCodes As Collection
    Set userCodes = users.GetEmployeeCodes(load.bumonCode)
    
    
    For Each userCode In userCodes
    
        '�]�ƈ����̎擾
        Dim userName As String
        userName = DataStorage.GetUserName(userCode)
        
        '�t�@�C���̒��o����������̐ݒ�
        userCodeFilter = fileProperty.UserCodeIdentifier & userCode & fileProperty.BreakIdentifier
        
        '�t�B���^�[�̎��s
        Dim filePathCollection As Collection
        Set filePathCollection = filter.AndFilter(bumonCodeFilter, targetDateFilter, userCodeFilter)
        
        '�t�B���^�[�̌��ʂ����݂���ꍇ
        If filePathCollection.Count > 0 Then
            '�t�@�C�����̎擾����
            fileProperty.InitFilePath data.SaveDirPath & "\" & filePathCollection(1)
            '�t�@�C���v���p�e�B�̕\��
            load.AddFileProperty userName, True, fileProperty.UpdatedDate
        Else
            '�t�@�C���v���p�e�B�̕\��
            load.AddFileProperty userName, False, Date
        End If
        
        
    Next userCode
    
End Sub

Sub LoadData()
    Dim load As New LoadSheetAccesser
    Dim data As New DataSheetAccesser
    Dim fileProperty As New FilePropertyManager
    Dim filter As New FileFilter
    
    load.ClearData
    
'''�������̎擾'''
    Dim bumonCodeFilter As String
    Dim targetDateFilter As String
    
    '�t�@�C���̒��o����������̐ݒ�
    bumonCodeFilter = fileProperty.BumonCodeIdentifier & load.bumonCode & fileProperty.BreakIdentifier
    targetDateFilter = fileProperty.DateIdentifier & Format(load.targetDate, "yyyymmdd") & fileProperty.BreakIdentifier
    
    '�t�B���^�[�̎��s
    Dim filePathCollection As Collection
    filter.DirPath = data.SaveDirPath
    Set filePathCollection = filter.AndFilter(bumonCodeFilter, targetDateFilter)
    
    
    For Each fileName In filePathCollection
        
        '�f�[�^�̎擾����
        data.InitSaveFileName CStr(fileName)
        data.InitOpenWorkBook
        data.InitWorkSheet
        '�t�@�C�����̎擾����
        fileProperty.InitFilePath data.SaveFilePath
        '���i�������
        load.WriteAllData data.GetAllData_NoHead
        '�f�[�^���[�N�u�b�N�����
        data.CloseWorkBook
        
    Next fileName
    
    '�O���[�v���ƏW�v�����ď�������
    Dim rs As ADODB.Recordset
    Set rs = load.AllGroupData
    load.ClearData
    load.WriteAllData RecordsetToCollection(rs)
    
    '�`�F�b�N�{�b�N�X��ǉ�
    '�J�n�s�ԍ�
    Dim rowIndex As Long
    rowIndex = load.DataStartRowIndex
    '�I���s�ԍ�
    Dim endRowIndex As Long
    endRowIndex = load.DataEndRowIndex
    
    '�󂯎��`�F�b�N�{�b�N�X
    Dim chkbox As Shape
    
    '�`�F�b�N�{�b�N�X
    Dim i As Long
    For i = rowIndex To endRowIndex
        '�`�F�b�N�{�b�N�X�̒ǉ�
        Set chkbox = load.AddCheckBox(i)
        '�`�F�b�N�{�b�N�X�ɃC�x���g�̕t�^
        chkbox.OnAction = "SaveOrderedData"
    Next i
    
    '�����t��������ݒ�
    load.ApplyConditionalFormatting
    
End Sub
