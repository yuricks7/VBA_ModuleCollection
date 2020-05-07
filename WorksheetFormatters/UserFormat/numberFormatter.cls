VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "NumberFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "数値の表示書式を定義します。"
Option Explicit

'━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

'正負
Private Enum ePositiveOrNegative
    Positive
    Negative
    Zero
End Enum

'連結単位
Public Enum eNotationType
    Positive
    Positive_Negative
    Positive_Negative_Zero
End Enum

'ユーザー書式で使える色
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

'半角スペース
Private pHalfWidthSpace As String

'━━━━━━━━━━━ コンストラクタ ━━━━━━━━━━━

Private Sub Class_Initialize()

    pHalfWidthSpace = " "

End Sub

'━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' 概要
'
' @param {String}            targetFormat          元になる表記
' @param {String}            unitName              数値の単位
' @param {eNotationType}     notationType          表記タイプ
' @param {eUserFormatColors} positiveColorConstant 正の数の色
' @param {String}            negativeSymbol        負の数に付ける記号
' @param {eUserFormatColors} negativeColorConstant 負の数の色
' @param {eUserFormatColors} zeroColorConstant     ゼロの色
'
' @return {String} 生成した文字列
'
Public Property Get NumberFormat(ByVal targetFormat As String, _
                                 Optional ByVal unitName As String = "", _
                                 Optional ByVal notationType As eNotationType = Positive_Negative, _
                                 Optional ByVal positiveColorConstant As eUserFormatColors = Unspecified, _
                                 Optional ByVal negativeSymbol As String = "▲", _
                                 Optional ByVal negativeColorConstant As eUserFormatColors = Red, _
                                 Optional ByVal zeroColorConstant As eUserFormatColors = Unspecified) As String

    '単位
    Dim unit As String: unit = addSpace(unitName)

    '正の数、負の数、ゼロの表記
    Dim positiveNumberNotaion As String, negativeNumberNotaion As String, zeroNotaion As String
    positiveNumberNotaion = getNumberNoation(ePositiveOrNegative.Positive, targetFormat, unit, positiveColorConstant)
    negativeNumberNotaion = getNumberNoation(ePositiveOrNegative.Negative, targetFormat, unit, negativeColorConstant, negativeSymbol)
              zeroNotaion = getNumberNoation(ePositiveOrNegative.Zero, "0", unit, zeroColorConstant)

    '連結
    NumberFormat = join(notationType, positiveNumberNotaion, negativeNumberNotaion, zeroNotaion)

End Property

''
' 読みやすいように、単位の後ろに空白を追加する
'
' @param {String} unitName 元の単位
'
' @return {String} 生成した文字列
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
' ユーザー書式を編集する
'
' @param {ePositiveOrNegative} aNumberNotation 内容
' @param {String}              sourceFormat    内容
' @param {String}              unit            内容
' @param {eUserFormatColors}   name            内容
' @param {String}              negativeSymbol  内容
'
' @return {String} 生成した文字列
'
Private Function getNumberNoation(ByVal aNumberNotation As ePositiveOrNegative, _
                                  Optional ByVal sourceFormat As String = "0", _
                                  Optional ByVal unit As String = "", _
                                  Optional ByVal colorConstant As eUserFormatColors = Unspecified, _
                                  Optional ByVal negativeSymbol As String = "▼") As String

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
' 連結して1つのユーザー書式とする
'
' @param {eNotationType} name 内容
' @param {String} positiveNumberNotaion 正の数
' @param {String} negativeNumberNotaion 負の数
' @param {String} zeroNotaion           ゼロ
'
' @return {String} 生成した文字列
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
' 色を選択する
'
' 【参考】
' ユーザー定義の表示形式で使える56色まとめ | Be COOL Users
' https://www.becoolusers.com/excel/custom-number-format-color.html
'
' @param {eUserFormatColors} formatColor ユーザー書式で使用する色
'
' @return {string} 内容
'
Public Function GetUserFormatColor(ByVal formatColor As eUserFormatColors) As String

    If formatColor = Unspecified Then
        GetUserFormatColor = ""
        Exit Function
    End If

    GetUserFormatColor = "[色" & formatColor & "]"

End Function


