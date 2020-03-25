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
' データの整形を実行
'
' @param {Range}       targetRange 変換の対象範囲
' @param {Long}        offsetWidth 出力先の相対列数
' @param {eOutputMode} outputMode  出力モード
'
Public Sub Run(ByVal targetRange As Range, _
               Optional ByVal offsetWidth As Long = 0, _
               Optional ByVal outputMode As eOutputMode)

    Select Case outputMode
        Case eOutputMode.eFixed
            '上書きされるので一旦避難
            Dim values As Variant: values = targetRange.Value
            
            Call UnifyZenkakuHankaku(targetRange, 0)
            Call RemoveLineBreaksInCell(targetRange, offsetWidth)
        
            'もとに戻す
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
' セル内改行を除去
'
' 【参考】
' セル内の改行を削除 | なんだ！カンタン！Excel塾
' https://kokodane.com/2013_macro_70.htm
'
' @param {Range} targetRange 変換の対象範囲
' @param {Long}  offsetWidth 出力先の相対列数
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
' 英数字を半角に、カタカナは全角に変換する
'
' 【参考】
' 「数字・アルファベットは半角・カタカナは全角」Excelで全角・半角を処理する関数・マクロ | EX-IT
' https://www.ex-it-blog.com/zenkaku-hankaku
'
' @param {Range} targetRange 変換の対象範囲
' @param {Long}  offsetWidth 出力先の相対列数
'
Public Sub UnifyZenkakuHankaku(ByRef targetRange As Range, _
                               Optional ByVal offsetWidth As Long = 0)

    Dim cell As Range
    For Each cell In targetRange
        'すべての語をいったん全角にする
        Dim allZenkakuChars As String
        allZenkakuChars = StrConv(cell.Value, vbWide)

        '1文字ずつチェックして置き換え
        Dim newChars As String: newChars = ""
        Dim i As Long
        For i = 1 To Len(allZenkakuChars)
            Dim char As String
            char = Mid(allZenkakuChars, i, 1)
            char = replaceWhenAlphanumerics(char)
            
            newChars = newChars & char
Continue:
        Next i
    
        '表記ゆれを統一
        newChars = replaceDoubleSpaces(newChars)
        newChars = replaceHyphens(newChars)

        '出力
        cell.Offset(0, offsetWidth).Value = newChars
    Next cell

End Sub

''
' 英数字を半角に置き換える
'
' 【参考】
' 「数字・アルファベットは半角・カタカナは全角」Excelで全角・半角を処理する関数・マクロ | EX-IT
' https://www.ex-it-blog.com/zenkaku-hankaku
'
' @param {String} char 文字
'
' @return {String} 修正後の文字列
'
Private Function replaceWhenAlphanumerics(ByVal char As String) As String
    
    '英数字のみ半角に置き換える
    If Not char Like "[ァ-ヶ]" Then char = StrConv(char, vbNarrow)

    replaceWhenAlphanumerics = char

End Function

''
' スペースが連なっているのを修正
'
' @param {String} chars 元の文字列
'
' @return {String} 修正後の文字列
'
Private Function replaceDoubleSpaces(ByVal chars As String) As String

    replaceDoubleSpaces = Replace(chars, "  ", " ")
    
End Function

''
' 検索用に横棒を統一
'
' @param {String} chars 元の文字列
'
' @return {String} 修正後の文字列
'
Private Function replaceHyphens(ByVal chars As String) As String

    Const HALF_WIDTH_HYPHEN As String = "-"
    Const FULL_WIDTH_HYPHEN As String = "ー"

    chars = Replace(chars, "‐", HALF_WIDTH_HYPHEN)
    chars = Replace(chars, "ー", FULL_WIDTH_HYPHEN)
    chars = Replace(chars, "ｰ", FULL_WIDTH_HYPHEN)

    replaceHyphens = chars
    
End Function
