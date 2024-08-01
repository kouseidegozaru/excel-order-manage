Attribute VB_Name = "Module1"
Sub test()

    ' クラスのインスタンスを作成
    Dim myClassInstance As New FilePropertyManager
    
    ' テスト用のファイルパスを設定
    Dim testFilePath As String
    testFilePath = "C:\Users\mfh077_user.MEFUREDMN\Desktop\excel-order-manage\発注管理\data\b40-u70-d20240725-.xlsx"
    
    ' filePath プロパティに値を設定
    myClassInstance.filePath = testFilePath
    
    ' 各プロパティの値を取得し、デバッグ出力
    Debug.Print "BumonCode: " & myClassInstance.BumonCode
    Debug.Print "UserCode: " & myClassInstance.UserCode
    Debug.Print "TargetDate: " & Format(myClassInstance.TargetDate, "yyyy-mm-dd")
    Debug.Print "UpdatedDate: " & Format(myClassInstance.UpdatedDate, "yyyy-mm-dd hh:nn:ss")

End Sub


