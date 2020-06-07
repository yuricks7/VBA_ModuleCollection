VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "FieldFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "���[�N�V�[�g�̗�𐮂��܂��B"
Option Explicit

'������������������������������������������������������ �v���V�[�W���p�̑I����

Public Enum eValueType
    eStandard
    eIdOrWords
    eNumbers
    eNumbers00
    eDatesOnly
    eDatesAndDaysOfTheWeek
    eDateTime
    eDateTimeAndDaysOfTheWeek
    [_eCounts]
End Enum

Public Enum eWidthType
    eGridPaper
    eMinutesAndHour
    eBoolean
    eOnlyWord
    eNarrow
    eDateWithDayOfTheWeek
    eDateTimeWithDayOfTheWeek
    eMiddleSentence
    eMax
    [_eCounts]
End Enum

'���������������������������������������������������������� �v���p�e�B�i�ϐ��j

Public Range As Range                  '�t�B�[���h�͈�
Public ValueType As eValueType         '�t�B�[���h�̌^
Public WidthType As eWidthType         '�񕝂̎��
Public CanAutoFit As Boolean           '�񕝂̎��������̉�
Public HorizontalAlignment As XlHAlign '�l�̐����ʒu

'�������������������������������������������������������������� �R���X�g���N�^

Private Sub Class_Initialize()

    ValueType = eIdOrWords

    WidthType = eOnlyWord
    CanAutoFit = False

    HorizontalAlignment = xlHAlignLeft

End Sub

'�������������������������������������������������� �v���p�e�B�i�v���V�[�W���j

''
' �Z���́u�\���`���v
'
' @param {eValueType} [aValueType] �t�B�[���h�Ɏg�p����l�̎��
'
' @return {String} �Z���́u�\���`���v�̐ݒ�l
'
Property Get DisplayFormat(Optional ByVal aValueType As eValueType = eValueType.[_eCounts]) As String

    Dim fieldTypes As Dictionary: Set fieldTypes = New Dictionary
    With fieldTypes
        .Add eValueType.eStandard, "G/�W��"
        .Add eValueType.eIdOrWords, "@"
        .Add eValueType.eNumbers, "#,##0_ ;[��]�� #,##0 "
        .Add eValueType.eNumbers00, "00"
        .Add eValueType.eDatesOnly, "yyyy/mm/dd"
        .Add eValueType.eDatesAndDaysOfTheWeek, "yyyy/mm/dd hh:mm"
        .Add eValueType.eDateTime, "yyyy/mm/dd (aaa)"
        .Add eValueType.eDateTimeAndDaysOfTheWeek, "yyyy/mm/dd (aaa) hh:mm"
    End With

    If aValueType = eValueType.[_eCounts] Then aValueType = ValueType
    DisplayFormat = fieldTypes.Item(aValueType)

End Property

''
' ��
'
' @param {eWidthType} [aWidthType] �񕝂̎��
'
' @return {Currency} ��
'
Property Get Width(Optional ByVal aWidthType As eWidthType = eWidthType.[_eCounts]) As Currency

    Dim widthDefinitions As Dictionary: Set widthDefinitions = New Dictionary
    With widthDefinitions
        .Add eWidthType.eGridPaper, 2.88
        .Add eWidthType.eMinutesAndHour, 3.88            ' 36px
        .Add eWidthType.eBoolean, 6.88                   ' 60px
        .Add eWidthType.eOnlyWord, 9.38                  ' 80px
        .Add eWidthType.eNarrow, 14.38                   '120px
        .Add eWidthType.eDateWithDayOfTheWeek, 14.38     '120px
        .Add eWidthType.eDateTimeWithDayOfTheWeek, 20.38 '168px
        .Add eWidthType.eMiddleSentence, 34.38           '280px
        .Add eWidthType.eMax, 46.88                      '380px

    End With

    If aWidthType = eWidthType.[_eCounts] Then aWidthType = WidthType
    Width = widthDefinitions.Item(aWidthType)

End Property

''
' �t�B�[���h��
'
' @return {String} ��
'
Property Get FieldName() As String

    FieldName = Me.Range.Cells(1, 1).Value

End Property

'�������������������������������������������������������������������� ���\�b�h

''
' ��̐ݒ��K�p����
'
Public Sub Apply()

    Dim field As Range: Set field = Me.Range

    With field
        .HorizontalAlignment = Me.HorizontalAlignment
        .NumberFormatLocal = Me.DisplayFormat

        .ColumnWidth = Me.Width

        If Me.CanAutoFit Then .AutoFit

    End With

End Sub

''
' �t�B�[���h���������l���擾���ēK�p����
'
Public Sub ApplyByHeader()

    Dim headerString As String: headerString = Me.FieldName
    Dim left2 As String:    left2 = Left(headerString, 2)
    Dim right1 As String: right1 = Right(headerString, 1)
    Dim right2 As String: right2 = Right(headerString, 2)
    Dim right3 As String: right3 = Right(headerString, 3)
    Dim right4 As String: right4 = Right(headerString, 4)

    With Me
        Select Case True '�������^�[���̂��߂Ɏg�p
            Case (left2 = "H_"), (left2 = "��_")
                .SetPropertiesForHour

            Case (left2 = "M_"), (left2 = "��_")
                .SetPropertiesForMinutes

            Case right2 = "����"
                .SetPropertiesForDateTime

            Case right1 = "��"
                .SetPropertiesForDate

            Case right2 = "ID"
                .SetPropertiesForId

            Case (right2 = "No"), (right3 = "No.")
                .SetPropertiesForRecordNumber

            Case (right2 = "�J�e"), (right4 = "�J�e�S��")
                .SetPropertiesForCategory

            Case (left2 = "is"), (right1 = "�H"), (right1 = "?")
                .SetPropertiesForBoolean

            Case Else
                .ValueType = eStandard
                .WidthType = eMiddleSentence
                .HorizontalAlignment = xlHAlignLeft

        End Select
    End With

    Me.Apply

End Sub

''
' ID
'
Public Sub SetPropertiesForId()

    With Me
        .ValueType = eIdOrWords
        .WidthType = eOnlyWord
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' �A�Ԃ�4�`6�����x�̐��l�Ȃ�
'
Public Sub SetPropertiesForRecordNumber()

    With Me
        .ValueType = eNumbers
        .WidthType = eOnlyWord
        .CanAutoFit = True
        .HorizontalAlignment = xlHAlignRight
    End With

End Sub

''
' ����
'
Public Sub SetPropertiesForDateTime()

    With Me
        .ValueType = eDateTimeAndDaysOfTheWeek
        .WidthType = eDateTimeWithDayOfTheWeek
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' ���t
'
Public Sub SetPropertiesForDate()

    With Me
        .ValueType = eDatesAndDaysOfTheWeek
        .WidthType = eDateWithDayOfTheWeek
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' ��
'
Public Sub SetPropertiesForHour()

    With Me
        .ValueType = eNumbers00
        .WidthType = eMinutesAndHour
        .HorizontalAlignment = xlHAlignRight
    End With

End Sub

''
' ��
'
Public Sub SetPropertiesForMinutes()

    With Me
        .ValueType = eNumbers00
        .WidthType = eMinutesAndHour
        .HorizontalAlignment = xlHAlignLeft
    End With

End Sub

''
' �J�e�S��
'
Public Sub SetPropertiesForCategory()

    With Me
        .ValueType = eStandard
        .WidthType = eOnlyWord
        .CanAutoFit = True
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' Is�`�n
'
Public Sub SetPropertiesForBoolean()

    With Me
        .ValueType = eStandard
        .WidthType = eBoolean
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub