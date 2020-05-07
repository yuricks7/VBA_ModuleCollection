VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataCleanser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eOutputMode
    eFixed
    eOneByOne
End Enum

''
' �f�[�^�̐��`�����s
'
' @param {Range}       targetRange �ϊ��̑Ώ۔͈�
' @param {Long}        offsetWidth �o�͐�̑��Η�
' @param {eOutputMode} outputMode  �o�̓��[�h
'
Public Sub Run(ByVal targetRange As Range, _
               Optional ByVal offsetWidth As Long = 0, _
               Optional ByVal outputMode As eOutputMode)

    Select Case outputMode
        Case eOutputMode.eFixed
            '�㏑�������̂ň�U���
            Dim values As Variant: values = targetRange.Value
            
            Call UnifyZenkakuHankaku(targetRange, 0)
            Call RemoveLineBreaksInCell(targetRange, offsetWidth)
        
            '���Ƃɖ߂�
            targetRange.Value = values
        
        Case eOutputMode.eOneByOne
            Call UnifyZenkakuHankaku( _
                targetRange, _
                offsetWidth _
            )
            
            Call RemoveLineBreaksInCell( _
                targetRange.Offset(0, offsetWidth), _
                1 _
            )
    End Select

End Sub

''
' �Z�������s������
'
' �y�Q�l�z
' �Z�����̉��s���폜 | �Ȃ񂾁I�J���^���IExcel�m
' https://kokodane.com/2013_macro_70.htm
'
' @param {Range} targetRange �ϊ��̑Ώ۔͈�
' @param {Long}  offsetWidth �o�͐�̑��Η�
'
Public Sub RemoveLineBreaksInCell(ByVal targetRange As Range, _
                                  Optional ByVal offsetWidth As Long = 0)

    Dim cell As Range
    Dim str As String
    For Each cell In targetRange
        If cell.Value = "" Then GoTo Continue
        
        cell.Offset(0, offsetWidth).Value = Replace(cell.Value, vbLf, "")
Continue:
    Next cell

End Sub

''
' �p�����𔼊p�ɁA�J�^�J�i�͑S�p�ɕϊ�����
'
' �y�Q�l�z
' �u�����E�A���t�@�x�b�g�͔��p�E�J�^�J�i�͑S�p�vExcel�őS�p�E���p����������֐��E�}�N�� | EX-IT
' https://www.ex-it-blog.com/zenkaku-hankaku
'
' @param {Range} targetRange �ϊ��̑Ώ۔͈�
' @param {Long}  offsetWidth �o�͐�̑��Η�
'
Public Sub UnifyZenkakuHankaku(ByRef targetRange As Range, _
                               Optional ByVal offsetWidth As Long = 0)

    Dim cell As Range
    For Each cell In targetRange
        '���ׂĂ̌����������S�p�ɂ���
        Dim allZenkakuChars As String
        allZenkakuChars = StrConv(cell.Value, vbWide)

        '1�������`�F�b�N���Ēu������
        Dim newChars As String: newChars = ""
        Dim i As Long
        For i = 1 To Len(allZenkakuChars)
            Dim char As String
            char = Mid(allZenkakuChars, i, 1)
            char = replaceWhenAlphanumerics(char)
            
            newChars = newChars & char
Continue:
        Next i
    
        '�\�L���𓝈�
        newChars = replaceDoubleSpaces(newChars)
        newChars = replaceHyphens(newChars)

        '�o��
        cell.Offset(0, offsetWidth).Value = newChars
    Next cell

End Sub

''
' �p�����𔼊p�ɒu��������
'
' �y�Q�l�z
' �u�����E�A���t�@�x�b�g�͔��p�E�J�^�J�i�͑S�p�vExcel�őS�p�E���p����������֐��E�}�N�� | EX-IT
' https://www.ex-it-blog.com/zenkaku-hankaku
'
' @param {String} char ����
'
' @return {String} �C����̕�����
'
Private Function replaceWhenAlphanumerics(ByVal char As String) As String
    
    '�p�����̂ݔ��p�ɒu��������
    If Not char Like "[�@-��]" Then char = StrConv(char, vbNarrow)

    replaceWhenAlphanumerics = char

End Function

''
' �X�y�[�X���A�Ȃ��Ă���̂��C��
'
' @param {String} chars ���̕�����
'
' @return {String} �C����̕�����
'
Private Function replaceDoubleSpaces(ByVal chars As String) As String

    replaceDoubleSpaces = Replace(chars, "  ", " ")
    
End Function

''
' �����p�ɉ��_�𓝈�
'
' @param {String} chars ���̕�����
'
' @return {String} �C����̕�����
'
Private Function replaceHyphens(ByVal chars As String) As String

    Const HALF_WIDTH_HYPHEN As String = "-"
    Const FULL_WIDTH_HYPHEN As String = "�["

    chars = Replace(chars, "�]", HALF_WIDTH_HYPHEN)
    chars = Replace(chars, "�[", FULL_WIDTH_HYPHEN)
    chars = Replace(chars, "�", FULL_WIDTH_HYPHEN)

    replaceHyphens = chars
    
End Function