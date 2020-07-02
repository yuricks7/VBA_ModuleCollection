VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = true
Attribute VB_Exposed         = False
Attribute VB_Description     = "�X�e�[�^�X�o�[�𐧌䂵�ă}�N���̐i������\�����܂��B"
Option Explicit

'���������������� ���W���[�����x���ϐ� ������������������

'�񋓑�
'Public Enum e�񋓑�
'    �����o�[ = 1
'End Enum

'�\����
'Public Type t�\����
'    �����o�[ As �^
'End Type

'���W���[�����x���ϐ�
Private pApps As Application
Private pMessage As String

'���������������� �v���p�e�B ������������������

Public Property Get Apps() As Application
Attribute Apps.VB_Description = "�X�e�[�^�X�o�[�Ŏg�p����Application�I�u�W�F�N�g"

    Set Apps = pApps

End Property
Public Property Set Apps(ByVal aApps As Application)
Attribute Apps.VB_Description = "�X�e�[�^�X�o�[�Ŏg�p����Application�I�u�W�F�N�g"

    Set pApps = aApps

End Property

''
' �f�t�H���g�ŕ\�����镶����
'
Public Property Get Message() As String
Attribute Message.VB_Description = "�f�t�H���g�ŕ\�����镶����"
Attribute Message.VB_UserMemId   = 0

    Message = pMessage

End Property
Public Property Let Message(ByVal msg As String)
Attribute Message.VB_Description = "�f�t�H���g�ŕ\�����镶����"
Attribute Message.VB_UserMemId   = 0

    pMessage = msg

End Property

'���������������� �R���X�g���N�^�� ������������������

''
' ����������
'
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "�N���X�̏�����"

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

''
' �I��������
'
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "�N���X�̌�n��"

'    �g���Ƃ��͐�ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."

End Sub

'���������������� ���\�b�h ������������������

''
' �X�e�[�^�X�o�[�Ƀ��b�Z�[�W��_�ł�����
'
' @param {String} msg �\�����镶����
'
Public Sub Flash(Optional ByVal msg As String)
Attribute Flash.VB_Description = "�X�e�[�^�X�o�[�Ƀ��b�Z�[�W��_�ł�����"

    If msg = "" Then msg = pMessage

    With pApps
        .StatusBar = False
        .StatusBar = msg
    End With

End Sub

''
' �X�e�[�^�X�o�[�Ƀ��b�Z�[�W��\������
'
' @param {String} msg �\�����镶����
'
Public Sub Display(Optional ByVal msg As String)
Attribute Display.VB_Description = "�X�e�[�^�X�o�[�Ƀ��b�Z�[�W��\������"

    If msg = "" Then msg = pMessage

    pApps.StatusBar = msg

End Sub

''
' �X�e�[�^�X�o�[�̕\������������
'
Public Sub Clear()
Attribute Clear.VB_Description = "�X�e�[�^�X�o�[�̕\������������"

    pApps.StatusBar = False

End Sub

''
' �X�e�[�^�X�o�[�ɐi����\������
'
' �y�Q�l�z
' Excel VBA ���Ԃ̂�����}�N���̐i���󋵂�\������ | MyRecord
' https://kosapi.com/post-2305/
'
' @param {Long}   currentCounts ���݂̌���
' @param {Long}   totalCounts   �����\�茏��
' @param {String} units         �P�ʕ\�L
' @param {String} description   �������e
'
Public Sub ShowProgressWithPercentage(ByVal currentCounts As Long, _
                                      ByVal totalCounts As Long, _
                                      Optional ByVal description As String _
                                                     = "�o��")
Attribute ShowProgressWithPercentage.VB_Description = "�X�e�[�^�X�o�[�ɐi����\������" _

    With Me
        '�O���t�\���p
        Dim currentSymbols As String
        Const MAX_COUNTS As Long = 10
        Dim blackCounts As Long
        blackCounts = Int((currentCounts / totalCounts) * MAX_COUNTS)
        currentSymbols = String(blackCounts, "��") & String(MAX_COUNTS, "��")

        '�X�e�[�^�X��
        Dim m As String
        m = description & "��..."
        m = m & " " & Left(currentSymbols, MAX_COUNTS)
        m = m & " " & Format(currentCounts / totalCounts, "0 %")

        .Display m
    End With

End Sub

''
' �X�e�[�^�X�o�[�ɐi����\������
'
' �y�Q�l�z
' Excel VBA ���Ԃ̂�����}�N���̐i���󋵂�\������ | MyRecord
' https://kosapi.com/post-2305/
'
' @param {Long}   currentCounts ���݂̌���
' @param {Long}   totalCounts   �����\�茏��
' @param {String} units         �P�ʕ\�L
' @param {String} description   �������e
'
Public Sub ShowProgressWithNumbers(ByVal currentCounts As Long, _
                                   ByVal totalCounts As Long, _
                                   Optional ByVal units As String = "��", _
                                   Optional ByVal description As String _
                                                  = "�o��")
Attribute ShowProgressWithNumbers.VB_Description = "�X�e�[�^�X�o�[�ɐi����\������"

    With Me

        '�O���t�\���p
        Dim currentSymbols As String
        Const MAX_COUNTS As Long = 10
        Dim blackCounts As Long
        blackCounts = Int((currentCounts / totalCounts) * MAX_COUNTS)
        currentSymbols = String(blackCounts, "��") & String(MAX_COUNTS, "��")

        '�X�e�[�^�X��
        Dim m As String
        m = description & "��..."
        m = m & " " & Left(currentSymbols, MAX_COUNTS)
        m = m & " " & Format(currentCounts, "#,##0") & " " & units
        m = m & " / " & Format(totalCounts, "#,##0") & " " & units

        .Display m
    End With

End Sub

''
' �X�e�[�^�X�o�[�ɃA�j���[�V������\������
'
' @param {Long} currentCounts ���݂̌���
' @param {Long} divisor       ����
'
Public Sub ShowWaitingAnime(ByVal currentCounts As Long, _
                            Optional ByVal msg As String _
                                           = "�f�[�^����͂��Ă��܂�", _
                            Optional ByVal divisor As Long = 1000)
Attribute ShowWaitingAnime.VB_Description = "�X�e�[�^�X�o�[�ɃA�j���[�V������\������"

    With Me
        '�A�j���[�V������
        Dim dotMark As String
        Dim criteria As Double: criteria = currentCounts / divisor
        Select Case criteria
            Case Is < 3
                dotMark = "  "
            Case Is < 6
                dotMark = ".  "
            Case Is < 9
                dotMark = ".. "
            Case Is >= 9
                dotMark = "..."
        End Select

        .Display m
    End With

End Sub