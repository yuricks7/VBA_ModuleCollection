VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Option Explicit

' ���������������������� ���W���[���ϐ� ����������������������

Private pRecord As Record
Private pRecords As Collection

Private pValues As Variant

' ���������������������� �R���X�g���N�^�� ����������������������

''
' ������
'
Private Sub Class_Initialize()

    Set pRecords = New Collection

End Sub

''
' �V�[�g�̒l����荞��
'
' @param {Worksheet} srcSheet ���[�N�V�[�g
'
' @return {Records} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Import(ByRef srcSheet As Worksheet) As Records

    Me.Values = srcSheet.UsedRange.Value
    Call importRowsData(srcSheet)

    Set Import = Me

End Function

''
' �V�[�g�̒l��Array1d�N���X�̃R���N�V�����ɕϊ�����
'
' @param {Worksheet} srcSheet ���[�N�V�[�g
'
Private Sub importRowsData(ByRef srcSheet As Worksheet)

    '�������ߖ�̂��߁A�g�p�ςݔ͈݂͂̂Ƃ���
    Dim importRange As Range
    With srcSheet.Cells
        Set importRange = .Range( _
                              .Rows(1), _
                              .Cells(.Rows.Count, 1).End(xlUp) _
                          )
    End With

    Dim i As Long
    For i = 1 To importRange.Rows.Count
        Dim r As Record: Set r = New Record
        r.Import (srcSheet.Rows(i))

        pRecords.Add r
    Next i

End Sub

' ���������������������� �v���p�e�B ����������������������

''
' �V�[�g�̒l�i�񎟌��z��j
'
Public Property Get Values() As Variant
    Values = pValues
End Property
Public Property Let Values(ByVal aValues As Variant)
    pValues = aValues
End Property

''
' ���R�[�h���擾����
'
' @return {Collection} �V�[�g1����
'
Property Get Items() As Collection
Attribute Items.VB_Description = "�V�[�g1�����̃��R�[�h���i�[�����R���N�V����"
Attribute Items.VB_UserMemId   = 0

    Set Items = pRecords

End Property

' ���������������������� ���\�b�h ����������������������

''
' ���R�[�h����1�s���̒l���擾����
'
' @param {Long} Item�̃L�[�i���l�j
'
' @return {Variant} ���R�[�h1�s��
'
Public Function Item(ByVal key As Long) As Record

    Set Item = pRecords.Item(key)

End Function