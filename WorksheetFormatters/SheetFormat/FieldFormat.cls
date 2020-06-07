VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "FieldFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "ワークシートの列を整えます。"
Option Explicit

'━━━━━━━━━━━━━━━━━━━━━━━━━━━ プロシージャ用の選択肢

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

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ プロパティ（変数）

Public Range      As Range                  'フィールド範囲
Public ValueType  As eValueType         'フィールドの型
Public WidthType  As eWidthType         '列幅の種類
Public CanAutoFit As Boolean           '列幅の自動調整の可否
Public HorizontalAlignment As XlHAlign '値の水平位置

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ コンストラクタ

Private Sub Class_Initialize()

    ValueType = eIdOrWords

    WidthType  = eOnlyWord
    CanAutoFit = False

    HorizontalAlignment = xlHAlignLeft

End Sub

'━━━━━━━━━━━━━━━━━━━━━━━━━ プロパティ（プロシージャ）

''
' セルの「表示形式」
'
' @param {eValueType} [aValueType] フィールドに使用する値の種類
'
' @return {String} セルの「表示形式」の設定値
'
Property Get DisplayFormat(Optional ByVal aValueType As eValueType = eValueType.[_eCounts]) As String

    Dim fieldTypes As Dictionary: Set fieldTypes = New Dictionary
    With fieldTypes
        .Add eValueType.eStandard,  "G/標準"
        .Add eValueType.eIdOrWords, "@"
        .Add eValueType.eNumbers,   "#,##0_ ;[赤]▼ #,##0 "
        .Add eValueType.eNumbers00, "00"
        .Add eValueType.eDatesOnly, "yyyy/mm/dd"
        .Add eValueType.eDatesAndDaysOfTheWeek,    "yyyy/mm/dd (aaa)"
        .Add eValueType.eDateTime,  "yyyy/mm/dd hh:mm"
        .Add eValueType.eDateTimeAndDaysOfTheWeek, "yyyy/mm/dd (aaa) hh:mm"
    End With

    If aValueType = eValueType.[_eCounts] Then aValueType = ValueType
    DisplayFormat = fieldTypes.Item(aValueType)

End Property

''
' 列幅
'
' @param {eWidthType} [aWidthType] 列幅の種類
'
' @return {Currency} 列幅
'
Property Get Width(Optional ByVal aWidthType As eWidthType = eWidthType.[_eCounts]) As Currency

    Dim widthDefinitions As Dictionary: Set widthDefinitions = New Dictionary
    With widthDefinitions
        .Add eWidthType.eGridPaper,       2.88
        .Add eWidthType.eMinutesAndHour,  3.88           ' 36px
        .Add eWidthType.eBoolean,         6.88           ' 60px
        .Add eWidthType.eOnlyWord,        9.38           ' 80px
        .Add eWidthType.eNarrow,         14.38           '120px
        .Add eWidthType.eDateWithDayOfTheWeek,     14.38 '120px
        .Add eWidthType.eDateTimeWithDayOfTheWeek, 20.38 '168px
        .Add eWidthType.eMiddleSentence, 34.38           '280px
        .Add eWidthType.eMax,            46.88                      '380px

    End With

    If aWidthType = eWidthType.[_eCounts] Then aWidthType = WidthType
    Width = widthDefinitions.Item(aWidthType)

End Property

''
' フィールド名
'
' @return {String} 列幅
'
Property Get FieldName() As String

    FieldName = Me.Range.Cells(1, 1).Value

End Property

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ メソッド

''
' 列の設定を適用する
'
Public Sub Apply()

    Dim field As Range: Set field = Me.Range

    With field
        .HorizontalAlignment = Me.HorizontalAlignment
        .NumberFormatLocal   = Me.DisplayFormat

        .ColumnWidth = Me.Width

        If Me.CanAutoFit Then .AutoFit

    End With

End Sub

''
' フィールド名から既定値を取得して適用する
'
Public Sub ApplyByHeader()

    Dim headerString As String: headerString = Me.FieldName
    Dim left2 As String:   left2 =  Left(headerString, 2)
    Dim right1 As String: right1 = Right(headerString, 1)
    Dim right2 As String: right2 = Right(headerString, 2)
    Dim right3 As String: right3 = Right(headerString, 3)
    Dim right4 As String: right4 = Right(headerString, 4)

    With Me
        Select Case True '早期リターンのために使用
            Case (left2 = "H_"), (left2 = "時_")
                .SetPropertiesForHour

            Case (left2 = "M_"), (left2 = "分_")
                .SetPropertiesForMinutes

            Case (left2 = "is"), (left2 = "Is"), (right1 = "？"), (right1 = "?")
                .SetPropertiesForBoolean

            Case right2 = "日時"
                .SetPropertiesForDateTime

            Case right1 = "日"
                .SetPropertiesForDate

            Case right2 = "ID"
                .SetPropertiesForId

            Case (right2 = "No"), (right3 = "No.")
                .SetPropertiesForRecordNumber

            Case (right2 = "カテ"), (right4 = "カテゴリ")
                .SetPropertiesForCategory

            Case Else
                .ValueType = eNumbers
                .WidthType = eMiddleSentence
                .CanAutoFit = True
                .HorizontalAlignment = xlHAlignRight

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

        If Me.Range.Width > .Width(eOnlyWord) Then
            .WidthType  = eNarrow
            .CanAutoFit = True

        Else
            .WidthType  = eOnlyWord
            .CanAutoFit = False

        End If

        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' 連番や4～6桁程度の数値など
'
Public Sub SetPropertiesForRecordNumber()

    With Me

        .ValueType = eNumbers

        If Me.Range.Width > .Width(eOnlyWord) Then
            .WidthType  = eNarrow
            .CanAutoFit = True

        Else
            .WidthType  = eOnlyWord
            .CanAutoFit = False

        End If

        .HorizontalAlignment = xlHAlignRight
    End With

End Sub

''
' 日時
'
Public Sub SetPropertiesForDateTime()

    With Me
        .ValueType = eDateTimeAndDaysOfTheWeek
        .WidthType = eDateTimeWithDayOfTheWeek
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' 日付
'
Public Sub SetPropertiesForDate()

    With Me
        .ValueType = eDatesAndDaysOfTheWeek
        .WidthType = eDateWithDayOfTheWeek
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub

''
' 時
'
Public Sub SetPropertiesForHour()

    With Me
        .ValueType = eNumbers00
        .WidthType = eMinutesAndHour
        .HorizontalAlignment = xlHAlignRight
    End With

End Sub

''
' 分
'
Public Sub SetPropertiesForMinutes()

    With Me
        .ValueType = eNumbers00
        .WidthType = eMinutesAndHour
        .HorizontalAlignment = xlHAlignLeft
    End With

End Sub

''
' カテゴリ
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
' Is～系
'
Public Sub SetPropertiesForBoolean()

    With Me
        .ValueType = eStandard
        .WidthType = eBoolean
        .HorizontalAlignment = xlHAlignCenter
    End With

End Sub
