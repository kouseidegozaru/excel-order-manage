Attribute VB_Name = "RockEvents"
'�ύX�������Ă΂��Ƃ��̕ύX�����m���Ă܂��ύX���N����A�������܂�Ă��܂��̂�
'��̕ύX�������I���܂Ō��m�𖳎�����

Public isIgnoreChange As Boolean

Sub SetIgnoreState(isIgnore As Boolean)
    isIgnoreChange = isIgnore
    
End Sub

Function GetIgnoreState() As Boolean
    GetIgnoreState = isIgnoreChange
    
End Function
