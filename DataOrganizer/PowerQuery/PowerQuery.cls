VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���������������� �l�̒�` ������������������

'�񋓑�
'Public Enum e�񋓑�
'
'End Enum

'�\����
'Public Enum t�\����
'
'End Enum

'���������������� ���W���[�����x���ϐ� ������������������

Private pQueryTable As QueryTable
Private pQueryTables As QueryTables

Private pTargetBook As Workbook

'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()


End Sub

''
' �R���X�g���N�^
'
' @param {workbook} aTargetBook �N�G�������s����u�b�N
'
' @return {PowerQuery} Me
'
Public Function Init(ByRef aTargetBook As Workbook) As PowerQuery

    Set pTargetBook = aTargetBook
    
    Set Init = Me

End Function

'�I��������
Private Sub Class_Terminate()
'    �g���Ƃ��͐�ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'���������������� �v���p�e�B ������������������

Property Get TargetBook() As Workbook: Set TargetBook = pTargetBook: End Property

''
' �u�b�N���́u�o�b�N�O���E���h�N�G���v�̉ۂ��܂Ƃ߂Đ؂�ւ���
'
' �y���Ӂz
' "True" �ɂ���ƁA�N�G���̓}�N���̏I����Ɏ��s�����B
' �i���s���̓t�@�C���ɐG��Ȃ��̂ŁA�ǂ����Ȃ� "False" �ɂ��ă}�N���̒��Ŏ��s���������ǂ������B�j
'
' @param {boolean} isTrue
'
Public Property Let CanRunBackgroundQueries(ByVal isTrue As Boolean)

    With Me.TargetBook
        Dim i As Long
        For i = 1 To .Connections.Count
          'Excludes PowerPivot and other connections
          If .Connections(i).Type = xlConnectionTypeOLEDB Then
            .Connections(i).OLEDBConnection.BackgroundQuery = isTrue
          End If
        Next i
    End With

End Property

'���������������� ���\�b�h ������������������

''
' �u���ׂčX�V�v�̎��s
'
Public Sub RefreshAll(Optional ByRef targetSheet As Worksheet)
    '���ׂčX�V
    Me.TargetBook.RefreshAll
    
    '�V�[�g�̌�n��
    If Not targetSheet Is Nothing Then Call adjustRowHeights(targetSheet)

End Sub

''
' PowerQuery�Őݒ肵���N�G���̂ݎ��s����
'
' @param {String}    queryName ���e
' @param {Worksheet} targetSheet ���e
'
Public Sub ImportData(ByVal queryName As String, _
                      Optional ByRef targetSheet As Worksheet)

    '�f�[�^�̍X�V
    On Error GoTo Catch:

    With pTargetBook
        .Connections(queryName).Refresh
    End With
    
    '�V�[�g�̌�n��
    If Not targetSheet Is Nothing Then Call adjustRowHeights(targetSheet)

    Exit Sub

Catch:

    MsgBox "�f�[�^�̃C���|�[�g���ɉ��炩�̖�肪����܂����B" & vbCrLf _
         & "�����𒆎~���܂��B", vbInformation

End Sub

''
' �V�[�g�̌��h���𒲐�����
'
' @param {Worksheet} targetSheet �Ώۂ̃��[�N�V�[�g
'
Private Sub adjustRowHeights(ByRef targetSheet As Worksheet)

    With targetSheet
        With .Rows
            .Font.Size = 8
            .RowHeight = 14.4 '24px
            .VerticalAlignment = xlBottom '������
        End With
        
        .Rows(1).Font.Size = 9 '�w�b�_�[
    End With

End Sub