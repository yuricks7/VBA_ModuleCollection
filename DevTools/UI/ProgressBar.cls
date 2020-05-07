VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
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

'���������������� �v���p�e�B ������������������

Public Property Get Apps() As Application:             Set Apps = pApps: End Property
Public Property Set Apps(ByVal aApps As Application): Set pApps = aApps: End Property

'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'�I��������
Private Sub Class_Terminate()
'    �g���Ƃ��͐�ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'���������������� ���\�b�h ������������������

''
' �X�e�[�^�X�o�[�Ƀ��b�Z�[�W��_�ł�����
'
' @param {String} msg �\�����镶����
'
Public Sub Flash(ByVal msg As String)

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
Public Sub ShowMessage(ByVal msg As String)

    pApps.StatusBar = msg

End Sub

''
' �X�e�[�^�X�o�[�̕\������������
'
Public Sub Clear()

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
                                      Optional ByVal description As String = "�o��")

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

        Apps.StatusBar = m
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
                                   Optional ByVal description As String = "�o��")

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

        Apps.StatusBar = m
    End With

End Sub

''
' �X�e�[�^�X�o�[�ɃA�j���[�V������\������
'
' @param {Long} currentCounts ���݂̌���
' @param {Long} divisor       ����
'
Public Sub ShowWaitingAnime(ByVal currentCounts As Long, _
                            Optional ByVal msg As String = "�f�[�^����͂��Ă��܂�", _
                            Optional ByVal divisor As Long = 1000)

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

        With pApps
            .StatusBar = msg & dotMark
        End With
    End With

End Sub