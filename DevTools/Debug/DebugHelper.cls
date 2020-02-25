VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pErr As ErrObject

'━━━━━━━━━━━ プロパティ ━━━━━━━━━━━


'━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' 値の型をイミディエイトウィンドウに出力する
'
' 【参考】
' VarType | Office TANAKA
' http://officetanaka.net/excel/vba/function/VarType.htm
'
' @param {Variant} something 型を見たいもの
'
Public Sub PrintVbaType(ByRef something As Variant)

    '定数を取得
    Dim VarTypeConstant As Long
    VarTypeConstant = VarType(something)

    '定数の値から型の内容を取得する
    Dim varTypes As Variant
    varTypes = getVarTypes(VarTypeConstant)

    '出力
    Debug.Print "名前：" & TypeName(something) & "型"
    Debug.Print "定数：" & "(" & Format(VarTypeConstant, "0") & ")" & varTypes(0)
    Debug.Print "詳細：" & varTypes(1)
    Debug.Print

End Sub

''
' 型の内容を取得する
'
' 【参考】
' VarType | Office TANAKA
' http://officetanaka.net/excel/vba/function/VarType.htm
'
' 【VB】VarType関数 | オープンリファレンス
' http://www.openreference.org/articles/view/460
'
' 【思案中】
' 構造体の方が読みやすく書けそうだけど、クラスでは使えないかも。
'
' @param {Long} varTypeConstant VBA上の型を表す組み込み定数
'
' @return {Variant} 配列（組み込み定数, 意味）
'
Private Function getVarTypes(ByVal VarTypeConstant As Long) As Variant
    
    Dim typeDescriptions As Variant
    Select Case VarTypeConstant
        Case 0
            typeDescriptions = Array("vbEmpty", "Empty値")

        Case 1
            typeDescriptions = Array("vbNull", "Null値")

        Case 2
            typeDescriptions = Array("vbInteger", "整数型")

        Case 3
            typeDescriptions = Array("vbLong", "長整数型")

        Case 4
            typeDescriptions = Array("vbSingle", "単精度浮動小数点数型")

        Case 5
            typeDescriptions = Array("vbDouble", "倍精度浮動小数点数型")

        Case 6
            typeDescriptions = Array("vbCurrency", "通貨型")

        Case 7
            typeDescriptions = Array("vbDate", "日付型")

        Case 8
            typeDescriptions = Array("vbString", "文字列型")

        Case 9
            typeDescriptions = Array("vbObject", "オブジェクト")

        Case 10
            typeDescriptions = Array("vbError", "エラー値")

        Case 11
            typeDescriptions = Array("vbBoolean", "ブール型")

        Case 12
            typeDescriptions = Array("vbVariant", "バリアント型配列")

        Case 13
            typeDescriptions = Array("vbDataObject", "非OLEオートメーションオブジェクト")

        Case 14
            typeDescriptions = Array("vbDecimal", "10進数型")

        Case 17
            typeDescriptions = Array("vbByte", "バイト型")

        Case 36
            typeDescriptions = Array("vbUserDefinedType", "バリアント型（含 ユーザー定義型）")

        '違い↓がよくわからない…
        Case 8204
            typeDescriptions = Array("vbArray", "配列+型番号")

        '違い↑がよくわからない…
        Case 8192
            typeDescriptions = Array("vbArray", "配列（バリアント型以外）")

        '（万が一）どれにも当てはまらない場合
        Case Else
            typeDescriptions = Array("????", "????")

    End Select

    getVarTypes = typeDescriptions

End Function

''
' エラーの詳細を伝える
'
' @param {ErrObject} errObj             発生したエラーオブジェクト
' @param {String}    processDescription 実行中の処理（の概要）
'
Public Sub showErrMsg(ByRef errObj As ErrObject, _
                      ByVal processDescription As String)
    
    Set pErr = errObj
    
    With pErr
        '情報が無いときの代替メッセージ
        Dim errorSouce As String:       errorSouce = .Source
        Dim errorDescription As String: errorDescription = .description
        If errorSouce = "" Then errorSouce = "なし"
        If errorDescription = "" Then errorDescription = "詳細不明です…"
        
        'メッセージを編集
        Dim errMsg As String
        errMsg = processDescription & "に何かエラーがあったみたい。" & vbCrLf
        errMsg = errMsg & vbCrLf
        errMsg = errMsg & "【概要】" & vbCrLf
        errMsg = errMsg & "エラー番号：" & .Number & vbCrLf
        errMsg = errMsg & "発生箇所：" & errorSouce & vbCrLf
        errMsg = errMsg & "【詳細】" & vbCrLf
        errMsg = errMsg & errorDescription
    End With

    MsgBox errMsg

End Sub
