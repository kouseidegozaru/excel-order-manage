Attribute VB_Name = "RockEvents"
'�ύX�������Ă΂��Ƃ��̕ύX�����m���Ă܂��ύX���N����A�������܂�Ă��܂��̂�
'��̕ύX�������I���܂Ō��m�𖳎�����

'���̃��W���[���ł̓X�e�[�^�X�݂̂��Ǘ�
Public isIgnoreChange As Boolean

Public Property Let IsIgnoreChangeEvents(isIgnore As Boolean)
    isIgnoreChange = isIgnore
End Property

Public Property Get IsIgnoreChangeEvents() As Boolean
    IsIgnoreChangeEvents = isIgnoreChange
End Property
