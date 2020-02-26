VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���������������������� ���W���[���ϐ� ����������������������

Private pErr As ErrObject

'���������������������� �v���p�e�B ����������������������


'���������������������� ���\�b�h ����������������������

''
' �l�̌^���C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���
'
' �y�Q�l�z
' VarType | Office TANAKA
' http://officetanaka.net/excel/vba/function/VarType.htm
'
' @param {Variant} something �^������������
'
Public Sub PrintVbaType(ByRef something As Variant)

    '�萔���擾
    Dim VarTypeConstant As Long
    VarTypeConstant = VarType(something)

    '�萔�̒l����^�̓��e���擾����
    Dim varTypes As Variant
    varTypes = getVarTypes(VarTypeConstant)

    '�o��
    Debug.Print "���O�F" & TypeName(something) & "�^"
    Debug.Print "�萔�F" & "(" & Format(VarTypeConstant, "0") & ")" & varTypes(0)
    Debug.Print "�ڍׁF" & varTypes(1)
    Debug.Print

End Sub

''
' �^�̓��e���擾����
'
' �y�Q�l�z
' VarType | Office TANAKA
' http://officetanaka.net/excel/vba/function/VarType.htm
'
' �yVB�zVarType�֐� | �I�[�v�����t�@�����X
' http://www.openreference.org/articles/view/460
'
' �y�v�Ē��z
' �\���̂̕����ǂ݂₷���������������ǁA�N���X�ł͎g���Ȃ������B
'
' @param {Long} varTypeConstant VBA��̌^��\���g�ݍ��ݒ萔
'
' @return {Variant} �z��i�g�ݍ��ݒ萔, �Ӗ��j
'
Private Function getVarTypes(ByVal VarTypeConstant As Long) As Variant
    
    Dim typeDescriptions As Variant
    Select Case VarTypeConstant
        Case 0
            typeDescriptions = Array("vbEmpty", "Empty�l")

        Case 1
            typeDescriptions = Array("vbNull", "Null�l")

        Case 2
            typeDescriptions = Array("vbInteger", "�����^")

        Case 3
            typeDescriptions = Array("vbLong", "�������^")

        Case 4
            typeDescriptions = Array("vbSingle", "�P���x���������_���^")

        Case 5
            typeDescriptions = Array("vbDouble", "�{���x���������_���^")

        Case 6
            typeDescriptions = Array("vbCurrency", "�ʉ݌^")

        Case 7
            typeDescriptions = Array("vbDate", "���t�^")

        Case 8
            typeDescriptions = Array("vbString", "������^")

        Case 9
            typeDescriptions = Array("vbObject", "�I�u�W�F�N�g")

        Case 10
            typeDescriptions = Array("vbError", "�G���[�l")

        Case 11
            typeDescriptions = Array("vbBoolean", "�u�[���^")

        Case 12
            typeDescriptions = Array("vbVariant", "�o���A���g�^�z��")

        Case 13
            typeDescriptions = Array("vbDataObject", "��OLE�I�[�g���[�V�����I�u�W�F�N�g")

        Case 14
            typeDescriptions = Array("vbDecimal", "10�i���^")

        Case 17
            typeDescriptions = Array("vbByte", "�o�C�g�^")

        Case 36
            typeDescriptions = Array("vbUserDefinedType", "�o���A���g�^�i�� ���[�U�[��`�^�j")

        '�Ⴂ�����悭�킩��Ȃ��c
        Case 8204
            typeDescriptions = Array("vbArray", "�z��+�^�ԍ�")

        '�Ⴂ�����悭�킩��Ȃ��c
        Case 8192
            typeDescriptions = Array("vbArray", "�z��i�o���A���g�^�ȊO�j")

        '�i������j�ǂ�ɂ����Ă͂܂�Ȃ��ꍇ
        Case Else
            typeDescriptions = Array("????", "????")

    End Select

    getVarTypes = typeDescriptions

End Function

''
' �G���[�̏ڍׂ�`����
'
' @param {ErrObject} errObj             ���������G���[�I�u�W�F�N�g
' @param {String}    processDescription ���s���̏����i�̊T�v�j
'
Public Sub showErrMsg(ByRef errObj As ErrObject, _
                      ByVal processDescription As String)
    
    Set pErr = errObj
    
    With pErr
        '��񂪖����Ƃ��̑�փ��b�Z�[�W
        Dim errorSouce As String:       errorSouce = .Source
        Dim errorDescription As String: errorDescription = .description
        If errorSouce = "" Then errorSouce = "�Ȃ�"
        If errorDescription = "" Then errorDescription = "�ڍוs���ł��c"
        
        '���b�Z�[�W��ҏW
        Dim errMsg As String
        errMsg = processDescription & "�ɉ����G���[���������݂����B" & vbCrLf
        errMsg = errMsg & vbCrLf
        errMsg = errMsg & "�y�T�v�z" & vbCrLf
        errMsg = errMsg & "�G���[�ԍ��F" & .Number & vbCrLf
        errMsg = errMsg & "�����ӏ��F" & errorSouce & vbCrLf
        errMsg = errMsg & "�y�ڍׁz" & vbCrLf
        errMsg = errMsg & errorDescription
    End With

    MsgBox errMsg

End Sub