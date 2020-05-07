VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "AutoColoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "条件付き書式を設定します。"
Option Explicit

'━━━━━━━━━ プロパティ ━━━━━━━━━━

Private pTargetSheet As Worksheet

Property Get targetSheet() As Worksheet
    Set targetSheet = pTargetSheet
End Property

Property Set targetSheet(ByRef aTargetSheet As Worksheet)
    Set targetSheet = aTargetSheet
End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()


'    Debug.Print Now & " " & "The Instance for 'IpTool' is initialized."

End Sub

'終了時処理
Private Sub Class_Terminate()

    ' このあと変数を破棄してしまうので、先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."

    Set pTargetSheet = Nothing

End Sub

'━━━━━━━━━━ メソッド ━━━━━━━━━━

Public Sub DeleteAllConditionsIn(ByRef targetRange As Range)

    ' 一旦すべて消す
    targetRange.FormatConditions.Delete

End Sub


''
' 条件付き書式を適用する
'
' @param {string} 適用条件式
' @param {boolean} 順位を先頭にしてよいか
'
Public Function Apply(ByVal expression As String, _
                      ByVal canSetFirst As Boolean) As FormatCondition

    With Me.ColoringRange
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=expression

        ' 第1位にセットするかどうか
        '（ここを消すと「.fontメソッドで失敗しました」とかで
        '条件付き書式のセットに失敗するみたい…）
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

            ' 反映させる条件付き書式の重複を許可
            .StopIfTrue = Me.CanStopIfTrue
    End With

    Set Apply = fc

End Function

''
' 指定範囲に数値があったら色付けする
'
' @param {long} 条件に使う数値
' @param {object} 色付けする範囲
'
Public Sub whenIsNumberIn()



End Sub

''
' 指定列に値があったら行ごと色付けする
'
' @param {array} 条件に使う値の配列
' @param {object} この条件付き書式を設定する範囲
' @param {long} 検索する列（指定範囲内の番号）
' @param {long} 条件を満たしたときの行の背景色（既定：淡いグレー）
' @param {long} 条件を満たしたときのフォントの色（既定：濃いグレー）
' @param {long} 条件を満たしたときに斜体にするか
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
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=targetExpression
        ' 第1位にセット
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

            ' 反映させる条件付き書式の重複を許可
            .StopIfTrue = False
        End With

    End With

End Sub

''
' ▼FormatCondition.Priority property (Excel)
' https://docs.microsoft.com/ja-jp/office/vba/api/excel.formatcondition.priority
'
' Returns or sets the priority value of the conditional formatting rule.
' The priority determines the order of evaluation
' when multiple conditional formatting rules exist on a worksheet.

Private Function joinFormulas(ByRef keyCell As Range, _
                              ByRef keyWords As Variant) As String

    ' 戻り値
    Dim retExpression As String
    retExpression = "" ' 初期化

    Dim iParam As Long
'Debug.Print "現在:"; iParam; retExpression

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
                ' 要素が1つだけならここで終わり
                If (paramCounts = 0) Then Exit For

            Case Is < paramCounts
                retExpression = retExpression & "," & temp

            ' 最後ならカッコを閉じて終了
            Case paramCounts
                retExpression = "OR(" & retExpression & "," & temp & ")"
        End Select

'Debug.Print "現在:"; iParam; retExpression

    Next

    retExpression = "=" & retExpression

'Debug.Print "最終:"; iParam; retExpression

    joinFormulas = retExpression

End Function


''
' 選択範囲の2行目に関数があれば色付けする
'
' @param {object} 色付けする範囲
' @param {long} 背景色
' @param {long} フォントの色
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
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' 第1位にセット
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

            ' 反映させる条件付き書式の重複を許可
            .StopIfTrue = False
        End With

    End With
End Sub



''
' 関数を満たしたら色付けする
'
' @param {string} 条件式
' @param {object} 色付けする範囲
' @param {long} 背景色
' @param {long} フォントの色
' @param {boolean} 太字にするか
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
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' 第1位にセット
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

            ' 反映させる条件付き書式の重複を許可
            .StopIfTrue = False
        End With

    End With
End Sub

''
' 指定範囲に空白があったら色をつける
'
' @param {range} 色付けする範囲
' @param {long} 指定範囲の何列目
' @param {boolean} 行の絶対参照を許可するか
' @param {boolean} 列の絶対参照を許可するか
' @param {long} 背景色
' @param {long} フォントの色
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
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=tempExpression
        ' 第1位にセット
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

            ' 反映させる条件付き書式の重複を許可
            .StopIfTrue = False
        End With

    End With

End Sub

''
' 指定範囲にエラーがあったら色をつける
'
' @param {object} 色付けする範囲
Public Sub whereErrorIsOn()



End Sub

''
' 値が変化するところで横方向に罫線を引く
'
' @param {object} この条件付き書式を設定する範囲
' @param {Long} 罫線の基準とする列番号（指定範囲のうちで何列目？）
' @param {Long} 罫線の色
' @param {Long} 罫線の位置
' @param {Long} 罫線のスタイル
' @param {Long} 罫線の太さ
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
        ' 条件付き書式を追加
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:=targetExpression
        ' 第1位にセット
        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        Const FIRST As Long = 1
        With .FormatConditions(FIRST)
            With .Borders(linePositionConstant)
                .LineStyle = xlContinuous
                .Weight = lineWeightConstant
                .Color = lineColor
            End With

        ' 反映させる条件付き書式の重複を許可
        .StopIfTrue = False

        End With
    End With

End Sub


'''
'' 値が変化するところで縦方向に罫線を引く
''
'' @param {object} 罫線の基準とする行
'Public Sub drawColumnBorder()
'
'
'
'End Sub

