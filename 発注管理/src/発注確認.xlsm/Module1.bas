Attribute VB_Name = "Module1"
Sub test()

    ' �N���X�̃C���X�^���X���쐬
    Dim myClassInstance As New FilePropertyManager
    
    ' �e�X�g�p�̃t�@�C���p�X��ݒ�
    Dim testFilePath As String
    testFilePath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\�����Ǘ�\data\b40-u70-d20240725-.xlsx"
    
    ' filePath �v���p�e�B�ɒl��ݒ�
    myClassInstance.filePath = testFilePath
    
    ' �e�v���p�e�B�̒l���擾���A�f�o�b�O�o��
    Debug.Print "BumonCode: " & myClassInstance.BumonCode
    Debug.Print "UserCode: " & myClassInstance.UserCode
    Debug.Print "TargetDate: " & Format(myClassInstance.TargetDate, "yyyy-mm-dd")
    Debug.Print "UpdatedDate: " & Format(myClassInstance.UpdatedDate, "yyyy-mm-dd hh:nn:ss")

End Sub


