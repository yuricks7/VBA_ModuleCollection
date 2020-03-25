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
        Dim errorDescription As String: errorDescription = .Description
        If errorSouce = "" Then errorSouce = "なし"
        If errorDescription = "" Then errorDescription = "（詳細不明）"
    
        Dim errorNumber As Long: errorNumber = .Number
    End With

    Dim m As String: m = ""
    If errorNumber <> 0 Then
        m = m & "申し訳ありません。" & vbCrLf
        m = m & "マクロの実行中にエラーが発生しました。" & vbCrLf
        m = m & vbCrLf
        m = m & "迅速な解決のため、可能でしたら" & vbCrLf
        m = m & "下記4点を添えてこのツールの管理者へご連絡ください。" & vbCrLf
        m = m & vbCrLf
        m = m & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        m = m & "1. Excelのバージョン" & vbCrLf
        m = m & "2. 実行したマクロ名（orボタン名）" & vbCrLf
        m = m & "3. 入力した値" & vbCrLf
        m = m & "4. この画面のスクリーンショット" & vbCrLf
        m = m & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        m = m & vbCrLf
    End If
        
    m = m & "─────────────────────────────" & vbCrLf
    m = m & "▼エラー内容" & vbCrLf
    m = m & "─────────────────────────────" & vbCrLf
    m = m & "【推定箇所】" & vbCrLf
    m = m & "プロジェクト：" & errorSouce & vbCrLf
    m = m & "処理内容　　：" & processDescription & vbCrLf
    m = m & vbCrLf
    m = m & "【エラー】" & vbCrLf
    m = m & "番号：" & errorNumber & vbCrLf
    m = m & "説明：" & errorDescription & vbCrLf
    m = m & "─────────────────────────────" & vbCrLf
    
    MsgBox m

End Sub
