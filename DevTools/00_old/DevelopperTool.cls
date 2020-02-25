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

    MsgBox "イミディエイトウィンドウに出力しました。", vbInformation

End Sub

Public Sub PrintVbaType(ByRef something As Variant)

    '▼OfficeTANAKA
    'http://officetanaka.net/excel/vba/function/VarType.htm
    Dim varTypeConstant As Long
    varTypeConstant = VarType(something)

    Dim varTypes As Variant
    varTypes = getVarTypes(varTypeConstant)

    Debug.Print "(" & Format(varTypeConstant, "0") & ") " & varTypes(0)
    Debug.Print "内容：" & varTypes(1)
    Debug.Print "型名：" & TypeName(something)
    Debug.Print

End Sub

Private Function getVarTypes(ByVal varTypeConstant As Long) As Variant

    '▼VarType | Office TANAKA
    'http://officetanaka.net/excel/vba/function/VarType.htm

    '▼【VB】VarType関数 | オープンリファレンス
    'http://www.openreference.org/articles/view/460

    Dim ret As Variant
    Select Case varTypeConstant
        Case 0
            ret = Array("vbEmpty", "Empty値")

        Case 1
            ret = Array("vbNull", "Null値")

        Case 2
            ret = Array("vbInteger", "整数型")

        Case 3
            ret = Array("vbLong", "長整数型")

        Case 4
            ret = Array("vbSingle", "単精度浮動小数点数型")

        Case 5
            ret = Array("vbDouble", "倍精度浮動小数点数型")

        Case 6
            ret = Array("vbCurrency", "通貨型")

        Case 7
            ret = Array("vbDate", "日付型")

        Case 8
            ret = Array("vbString", "文字列型")

        Case 9
            ret = Array("vbObject", "オブジェクト")

        Case 10
            ret = Array("vbError", "エラー値")

        Case 11
            ret = Array("vbBoolean", "ブール型")

        Case 12
            ret = Array("vbVariant", "バリアント型配列")

        Case 13
            ret = Array("vbDataObject", "非OLEオートメーションオブジェクト")

        Case 14
            ret = Array("vbDecimal", "10進数型")

        Case 17
            ret = Array("vbByte", "バイト型")

        Case 36
            ret = Array("vbUserDefinedType", "バリアント型（含 ユーザー定義型）")

        '違いがよくわからない…
        Case 8204
            ret = Array("vbArray", "配列+型番号")

        Case 8192
            ret = Array("vbArray", "配列（バリアント型以外）")

        'どれにも当てはまらない場合
        Case Else
            ret = Array("????", "????")

    End Select

    getVarTypes = ret

End Function
