VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoColoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'������������������ �v���p�e�B ��������������������

Private pTargetSheet As Worksheet

Property Get targetSheet() As Worksheet
    Set targetSheet = pTargetSheet
End Property

Property Set targetSheet(ByRef aTargetSheet As Worksheet)
    Set targetSheet = aTargetSheet
End Property

'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()


'    Debug.Print Now & " " & "The Instance for 'IpTool' is initialized."

End Sub

'�I��������
Private Sub Class_Terminate()

    ' ���̂��ƕϐ���j�����Ă��܂��̂ŁA��ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."

    Set pTargetSheet = Nothing

End Sub

'�������������������� ���\�b�h ��������������������

Public Sub DeleteAllConditionsIn(ByRef targetRange As Range)

    ' ��U���ׂď���
    targetRange.FormatConditions.Delete

End Sub


''
' �����t��������K�p����
'
' @param {string} �K�p������
' @param {boolean} ���ʂ�擪�ɂ��Ă悢��
'
Public Function Apply(ByVal expression As String, _
                      ByVal canSetFirst As Boolean) As FormatCondition

    With Me.ColoringRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=expression

        ' ��1�ʂɃZ�b�g���邩�ǂ���
        '�i�����������Ɓu.font���\�b�h�Ŏ��s���܂����v�Ƃ���
        '�����t�������̃Z�b�g�Ɏ��s����݂����c�j
        If canSetFirst = True Then
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
        Else
            .FormatConditions(.FormatConditions.Count).SetLastPriority
        End If

        Const FIRST As Long = 1
        Dim fc As FormatCondition
        Set fc = .FormatConditions(FIRST)
    End With

    With fc
        .Interior.Color = Me.interiorColor

        With .Font
            .Color = Me.fontColor
            .Italic = Me.IsItalicWhenTrue
            .Bold = Me.IsBoldWhenTrue
        End With

            ' ���f����������t�������̏d��������
            .StopIfTrue = Me.CanStopIfTrue
    End With

    Set Apply = fc

End Function

''
' �w��͈͂ɐ��l����������F�t������
'
' @param {long} �����Ɏg�����l
' @param {object} �F�t������͈�
'
Public Sub whenIsNumberIn()



End Sub

''
' �w���ɒl����������s���ƐF�t������
'
' @param {array} �����Ɏg���l�̔z��
' @param {object} ���̏����t��������ݒ肷��͈�
' @param {long} ���������i�w��͈͓��̔ԍ��j
' @param {long} �����𖞂������Ƃ��̍s�̔w�i�F�i����F�W���O���[�j
' @param {long} �����𖞂������Ƃ��̃t�H���g�̐F�i����F�Z���O���[�j
' @param {long} �����𖞂������Ƃ��ɎΑ̂ɂ��邩
'
Public Sub setRowColors(ByRef searchValues As Variant, _
                        ByRef ColoringRange As Range, _
                        Optional ByVal searchColNumber As Long = 1, _
                        Optional ByVal interiorColor As Long, _
                        Optional ByVal fontColor As Long, _
                        Optional ByVal isItalic As Long = True, _
                        Optional ByVal isBold As Long = False)
'    With Me
'        If interiorColor = 0 Then interiorColor = .LightGray
'        If fontColor = 0 Then fontColor = .DarkGray
'    End With

    Dim keyCell As Range
    Set keyCell = ColoringRange.Cells(1, searchColNumber)
'Debug.Print keyCell.Address, keyCell.value, searchValues

    Dim targetExpression As String
    targetExpression = joinFormulas(keyCell, searchValues)

    With ColoringRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=targetExpression
        ' ��1�ʂɃZ�b�g
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)

            With .Font
                .Italic = isItalic
                .Color = fontColor
            End With

            If interiorColor = xlNone Then

            Else
                .Interior.Color = interiorColor

            End If

            ' ���f����������t�������̏d��������
            .StopIfTrue = False
        End With

    End With

End Sub

''
' ��FormatCondition.Priority property (Excel)
' https://docs.microsoft.com/ja-jp/office/vba/api/excel.formatcondition.priority
'
' Returns or sets the priority value of the conditional formatting rule.
' The priority determines the order of evaluation
' when multiple conditional formatting rules exist on a worksheet.

Private Function joinFormulas(ByRef keyCell As Range, _
                              ByRef keyWords As Variant) As String

    ' �߂�l
    Dim retExpression As String
    retExpression = "" ' ������

    Dim iParam As Long
'Debug.Print "����:"; iParam; retExpression

    Dim paramCounts As Long
    paramCounts = UBound(keyWords)
'Debug.Print "paramCounts:"; paramCounts

    For iParam = 0 To paramCounts
        Dim colAddress As String
        colAddress = keyCell.Address(RowAbsolute:=False, _
                                  ColumnAbsolute:=True)

        Dim keyword As String
        Const DOUBLE_QUATATION As String = """"
        If VarType(keyWords(iParam)) = vbString Then
            keyword = DOUBLE_QUATATION & keyWords(iParam) & DOUBLE_QUATATION

        Else
            keyword = keyWords(iParam)

        End If

        Dim temp As String
        temp = ""
        temp = colAddress & "=" & keyword

'        temp = colAddress & "=" & DOUBLE_QUATATION & keyWord & DOUBLE_QUATATION

        Select Case iParam
            Case 0
                retExpression = temp
                ' �v�f��1�����Ȃ炱���ŏI���
                If (paramCounts = 0) Then Exit For

            Case Is < paramCounts
                retExpression = retExpression & "," & temp

            ' �Ō�Ȃ�J�b�R����ďI��
            Case paramCounts
                retExpression = "OR(" & retExpression & "," & temp & ")"
        End Select

'Debug.Print "����:"; iParam; retExpression

    Next

    retExpression = "=" & retExpression

'Debug.Print "�ŏI:"; iParam; retExpression

    joinFormulas = retExpression

End Function


''
' �I��͈͂�2�s�ڂɊ֐�������ΐF�t������
'
' @param {object} �F�t������͈�
' @param {long} �w�i�F
' @param {long} �t�H���g�̐F
'
Public Sub hasFormula(ByRef ColoringRange As Range, _
                      ByVal interiorColor As Long, _
                      ByVal fontColor As Long)

    Dim keyCell As Range
    Dim keyAddress As String
    Set keyCell = ColoringRange.Cells(2, 1)
    keyAddress = keyCell.Address(RowAbsolute:=True, _
                              ColumnAbsolute:=False)

    Dim tempExpression As String
    tempExpression = "=isFormula(" & keyAddress & ")"

    With ColoringRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' ��1�ʂɃZ�b�g
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)

            With .Font
                .Bold = True
                .Color = fontColor
            End With

            If interiorColor = xlNone Then

            Else
                .Interior.Color = interiorColor

            End If

            ' ���f����������t�������̏d��������
            .StopIfTrue = False
        End With

    End With
End Sub



''
' �֐��𖞂�������F�t������
'
' @param {string} ������
' @param {object} �F�t������͈�
' @param {long} �w�i�F
' @param {long} �t�H���g�̐F
' @param {boolean} �����ɂ��邩
'
Public Sub IsSatisfied(ByVal condition As String, _
                       ByRef ColoringRange As Range, _
                       Optional ByVal fontColor As Long = xlNone, _
                       Optional ByVal interiorColor As Long = xlNone, _
                       Optional ByVal isBold As Boolean = False, _
                       Optional ByVal isItalic As Boolean = False)

    Dim tempExpression As String
    tempExpression = "=" & condition

    With ColoringRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' ��1�ʂɃZ�b�g
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)

            With .Font
                .Bold = isBold
                .Italic = isItalic
                .Color = fontColor
            End With

            If interiorColor = xlNone Then

            Else
                .Interior.Color = interiorColor

            End If

            ' ���f����������t�������̏d��������
            .StopIfTrue = False
        End With

    End With
End Sub

''
' �w��͈͂ɋ󔒂���������F������
'
' @param {range} �F�t������͈�
' @param {long} �w��͈͂̉����
' @param {boolean} �s�̐�ΎQ�Ƃ������邩
' @param {boolean} ��̐�ΎQ�Ƃ������邩
' @param {long} �w�i�F
' @param {long} �t�H���g�̐F
'
Public Sub hasBlank(ByRef ColoringRange As Range, _
                    ByVal creteriaColNum As Long, _
                    ByVal isAbsoluteRow As Boolean, _
                    ByVal isAbsoluteCol As Boolean, _
                    ByVal interiorColor As Long, _
                    ByVal fontColor As Long)

    Dim keyCell As Range
    Dim keyAddress As String
    Set keyCell = ColoringRange.Cells(1, creteriaColNum)
    keyAddress = keyCell.Address(RowAbsolute:=isAbsoluteRow, _
                              ColumnAbsolute:=isAbsoluteCol)

    Dim tempExpression As String
    tempExpression = "=isBlank(" & keyAddress & ")"

    With ColoringRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' ��1�ʂɃZ�b�g
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)

            With .Font
                .Bold = False
                .Color = fontColor
            End With

            If interiorColor = xlNone Then

            Else
                .Interior.Color = interiorColor

            End If

            ' ���f����������t�������̏d��������
            .StopIfTrue = False
        End With

    End With

End Sub

''
' �w��͈͂ɃG���[����������F������
'
' @param {object} �F�t������͈�
Public Sub whereErrorIsOn()



End Sub

''
' �l���ω�����Ƃ���ŉ������Ɍr��������
'
' @param {object} ���̏����t��������ݒ肷��͈�
' @param {Long} �r���̊�Ƃ����ԍ��i�w��͈͂̂����ŉ���ځH�j
' @param {Long} �r���̐F
' @param {Long} �r���̈ʒu
' @param {Long} �r���̃X�^�C��
' @param {Long} �r���̑���
'
Public Sub drawRowBorder(ByRef drawingRange As Range, _
                         Optional ByVal criteriaColNumber As Long = 1, _
                         Optional ByVal lineColor As Long, _
                         Optional ByVal linePositionConstant As Long = xlBottom, _
                         Optional ByVal lineStyleConstant As Long = xlContinuous, _
                         Optional ByVal lineWeightConstant As Long = xlThin)
'    If lineColor = 0 Then lineColor = Me.Black

    With drawingRange
        Dim upperAddress As String
        upperAddress = .Cells(1, criteriaColNumber).Address(RowAbsolute:=False, _
                                                         ColumnAbsolute:=True)
        Dim lowerAddress As String
        lowerAddress = .Cells(2, criteriaColNumber).Address(RowAbsolute:=False, _
                                                         ColumnAbsolute:=True)
    End With

    Dim targetExpression As String
    targetExpression = "=" & upperAddress & "<>" & lowerAddress

' Debug.Print targetExpression

    With drawingRange
        ' �����t��������ǉ�
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=targetExpression
        ' ��1�ʂɃZ�b�g
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)
            With .Borders(linePositionConstant)
                .LineStyle = xlContinuous
                .Weight = lineWeightConstant
                .Color = lineColor
            End With

        ' ���f����������t�������̏d��������
        .StopIfTrue = False

        End With
    End With

End Sub


'''
'' �l���ω�����Ƃ���ŏc�����Ɍr��������
''
'' @param {object} �r���̊�Ƃ���s
'Public Sub drawColumnBorder()
'
'
'
'End Sub
