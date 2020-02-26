VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DevelopperTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PrintHeaders(ByRef list As ListObject)

    Dim headerRange As Range
    Set headerRange = list.HeaderRowRange

    Dim listCounts As Long
    listCounts = list.ListColumns.Count

    Dim headers As Variant
    ReDim headers(listCounts)

    Dim i As Long
    For i = 1 To listCounts
        headers(i) = headerRange(1, i).Value
        Debug.Print headers(i)
    Next

    MsgBox "�C�~�f�B�G�C�g�E�B���h�E�ɏo�͂��܂����B", vbInformation

End Sub

Public Sub PrintVbaType(ByRef something As Variant)

    '��OfficeTANAKA
    'http://officetanaka.net/excel/vba/function/VarType.htm
    Dim varTypeConstant As Long
    varTypeConstant = VarType(something)

    Dim varTypes As Variant
    varTypes = getVarTypes(varTypeConstant)

    Debug.Print "(" & Format(varTypeConstant, "0") & ") " & varTypes(0)
    Debug.Print "���e�F" & varTypes(1)
    Debug.Print "�^���F" & TypeName(something)
    Debug.Print

End Sub

Private Function getVarTypes(ByVal varTypeConstant As Long) As Variant

    '��VarType | Office TANAKA
    'http://officetanaka.net/excel/vba/function/VarType.htm

    '���yVB�zVarType�֐� | �I�[�v�����t�@�����X
    'http://www.openreference.org/articles/view/460

    Dim ret As Variant
    Select Case varTypeConstant
        Case 0
            ret = Array("vbEmpty", "Empty�l")

        Case 1
            ret = Array("vbNull", "Null�l")

        Case 2
            ret = Array("vbInteger", "�����^")

        Case 3
            ret = Array("vbLong", "�������^")

        Case 4
            ret = Array("vbSingle", "�P���x���������_���^")

        Case 5
            ret = Array("vbDouble", "�{���x���������_���^")

        Case 6
            ret = Array("vbCurrency", "�ʉ݌^")

        Case 7
            ret = Array("vbDate", "���t�^")

        Case 8
            ret = Array("vbString", "������^")

        Case 9
            ret = Array("vbObject", "�I�u�W�F�N�g")

        Case 10
            ret = Array("vbError", "�G���[�l")

        Case 11
            ret = Array("vbBoolean", "�u�[���^")

        Case 12
            ret = Array("vbVariant", "�o���A���g�^�z��")

        Case 13
            ret = Array("vbDataObject", "��OLE�I�[�g���[�V�����I�u�W�F�N�g")

        Case 14
            ret = Array("vbDecimal", "10�i���^")

        Case 17
            ret = Array("vbByte", "�o�C�g�^")

        Case 36
            ret = Array("vbUserDefinedType", "�o���A���g�^�i�� ���[�U�[��`�^�j")

        '�Ⴂ���悭�킩��Ȃ��c
        Case 8204
            ret = Array("vbArray", "�z��+�^�ԍ�")

        Case 8192
            ret = Array("vbArray", "�z��i�o���A���g�^�ȊO�j")

        '�ǂ�ɂ����Ă͂܂�Ȃ��ꍇ
        Case Else
            ret = Array("????", "????")

    End Select

    getVarTypes = ret

End Function