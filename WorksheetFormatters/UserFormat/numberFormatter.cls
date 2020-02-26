VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "NumberFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Option Explicit

'���������������������� ���W���[���ϐ� ����������������������

'����
Private Enum ePositiveOrNegative
    Positive
    Negative
    Zero
End Enum

'�A���P��
Public Enum eNotationType
    Positive
    Positive_Negative
    Positive_Negative_Zero
End Enum

'���[�U�[�����Ŏg����F
Public Enum eUserFormatColors
    Unspecified = 0
    Black
    White
    Red
    LightGreen
    Blue
    Yellow
    Pink
    LightBlue
    Brown
    MosSreen
    DarkBlue
    LightBrown
    Purple
    BlueGreen
    LightGray
    DarkGray
    BluePurple
    RedPurple
    PaleYellow
    SeaBlue
    DarkPurple
    LightOrange
    JapaneseBlue
    LighterPurple
    Navy
    PinkPurple
    BrightYellow
    BrightSeaBlue
    DarkRedPurple
    DarkBrown
    DarkBlueGreen
    BrightBlue
    DarkSeaBlue
    LighterBlue
    LighterGreen
    LighterYellow
    LighterSeaBlue
    LightPink
    LightPurple
    PaleOrange
    PrettyBlue
    LightBlueGreen
    LightMosGreen
    DarkYellow
    Orange
    BrightOrange
    BlueGray
    Gray
    DarkNavy
    PrettyGreen
    BlackGreen
    BrownGray
    BrightBrown
    DarkPrettyBlue
    LightNavy
    BlueBlack
End Enum

'���p�X�y�[�X
Private pHalfWidthSpace As String

'���������������������� �R���X�g���N�^ ����������������������

Private Sub Class_Initialize()

    pHalfWidthSpace = " "

End Sub

'���������������������� �v���p�e�B ����������������������

''
' �T�v
'
' @param {String}            targetFormat          ���ɂȂ�\�L
' @param {String}            unitName              ���l�̒P��
' @param {eNotationType}     notationType          �\�L�^�C�v
' @param {eUserFormatColors} positiveColorConstant ���̐��̐F
' @param {String}            negativeSymbol        ���̐��ɕt����L��
' @param {eUserFormatColors} negativeColorConstant ���̐��̐F
' @param {eUserFormatColors} zeroColorConstant     �[���̐F
'
' @return {String} ��������������
'
Public Property Get NumberFormat(ByVal targetFormat As String, _
                                 Optional ByVal unitName As String = "", _
                                 Optional ByVal notationType As eNotationType = Positive_Negative, _
                                 Optional ByVal positiveColorConstant As eUserFormatColors = Unspecified, _
                                 Optional ByVal negativeSymbol As String = "��", _
                                 Optional ByVal negativeColorConstant As eUserFormatColors = Red, _
                                 Optional ByVal zeroColorConstant As eUserFormatColors = Unspecified) As String

    '�P��
    Dim unit As String: unit = addSpace(unitName)
    
    '���̐��A���̐��A�[���̕\�L
    Dim positiveNumberNotaion As String, negativeNumberNotaion As String, zeroNotaion As String
    positiveNumberNotaion = getNumberNoation(ePositiveOrNegative.Positive, targetFormat, unit, positiveColorConstant)
    negativeNumberNotaion = getNumberNoation(ePositiveOrNegative.Negative, targetFormat, unit, negativeColorConstant, negativeSymbol)
              zeroNotaion = getNumberNoation(ePositiveOrNegative.Zero, "0", unit, zeroColorConstant)
    
    '�A��
    NumberFormat = join(notationType, positiveNumberNotaion, negativeNumberNotaion, zeroNotaion)

End Property

''
' �ǂ݂₷���悤�ɁA�P�ʂ̌��ɋ󔒂�ǉ�����
'
' @param {String} unitName ���̒P��
'
' @return {String} ��������������
'
Private Function addSpace(Optional ByVal unitName As String = "") As String

    Dim unit As String
    If unitName = "" Then
        unit = ""
    Else
        unit = unitName & pHalfWidthSpace
    End If

    addSpace = unit

End Function

''
' ���[�U�[������ҏW����
'
' @param {ePositiveOrNegative} aNumberNotation ���e
' @param {String}              sourceFormat    ���e
' @param {String}              unit            ���e
' @param {eUserFormatColors}   name            ���e
' @param {String}              negativeSymbol  ���e
'
' @return {String} ��������������
'
Private Function getNumberNoation(ByVal aNumberNotation As ePositiveOrNegative, _
                                  Optional ByVal sourceFormat As String = "0", _
                                  Optional ByVal unit As String = "", _
                                  Optional ByVal colorConstant As eUserFormatColors = Unspecified, _
                                  Optional ByVal negativeSymbol As String = "��") As String

    Dim ret As String
    
    Select Case aNumberNotation
        Case ePositiveOrNegative.Positive
            ret = GetUserFormatColor(colorConstant) & sourceFormat & "_" & pHalfWidthSpace & unit
        
        Case ePositiveOrNegative.Negative
            ret = ";" & GetUserFormatColor(colorConstant) & negativeSymbol & sourceFormat & pHalfWidthSpace & unit
        
        Case ePositiveOrNegative.Zero
            ret = ";" & GetUserFormatColor(colorConstant) & sourceFormat & pHalfWidthSpace & unit
    
    End Select
    
    getNumberNoation = ret

End Function

''
' �A������1�̃��[�U�[�����Ƃ���
'
' @param {eNotationType} name ���e
' @param {String} positiveNumberNotaion ���̐�
' @param {String} negativeNumberNotaion ���̐�
' @param {String} zeroNotaion           �[��
'
' @return {String} ��������������
'
Private Function join(ByVal notationType As eNotationType, _
                      ByVal positiveNumberNotaion As String, _
                      Optional ByVal negativeNumberNotaion As String, _
                      Optional ByVal zeroNotaion As String) As String

    Dim ret As String
    
    Select Case notationType
        Case eNotationType.Positive
            ret = positiveNumberNotaion
        
        Case eNotationType.Positive_Negative
            ret = positiveNumberNotaion & negativeNumberNotaion
        
        Case eNotationType.Positive_Negative_Zero
            ret = positiveNumberNotaion & negativeNumberNotaion & zeroNotaion
    End Select

    join = ret

End Function

''
' �F��I������
'
' �y�Q�l�z
' ���[�U�[��`�̕\���`���Ŏg����56�F�܂Ƃ� | Be COOL Users
' https://www.becoolusers.com/excel/custom-number-format-color.html
'
' @param {eUserFormatColors} formatColor ���[�U�[�����Ŏg�p����F
'
' @return {string} ���e
'
Public Function GetUserFormatColor(ByVal formatColor As eUserFormatColors) As String

    If formatColor = Unspecified Then
        GetUserFormatColor = ""
        Exit Function
    End If
    
    GetUserFormatColor = "[�F" & formatColor & "]"

End Function

