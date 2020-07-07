VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Array1d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "一次元配列を扱いやすくします。"
Option Explicit

' ━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pArray As Variant

' ━━━━━━━━━━━ コンストラクタ等 ━━━━━━━━━━━

''
' コンストラクタの代わり
'
' @param {Long} [maxIndex] インデックスの最大値
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Init(Optional ByVal maxIndex As Long = 0) As Array1d
Attribute Init.VB_Description = "モジュール変数pArrayを初期化する"

    ReDim pArray(maxIndex) '配列の初期化

    Set Init = Me

End Function

''
' コンストラクタの代わり（Split関数の代用）
'
' @param {String}          srcString     元の文字列
' @param {String}          splitKey      検索キー
' @param {Long}            [lastIndex]   最大の分割数
' @param {VbCompareMethod} [compareMode] 文字列の検索条件
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Create(ByVal srcString As String, _
                       ByVal splitKey As String, _
                       Optional ByVal lastIndex As Long = -1, _
                       Optional ByVal compareMode As VbCompareMethod _
                                                   = vbBinaryCompare _
                       ) As Array1d
Attribute Create.VB_Description = "Split関数を使用し、文字列を配列化する"

    Dim tempArr As Variant
    tempArr = Split(srcString, splitKey, lastIndex, compareMode)

    Me.Init (UBound(tempArr))
    pArray = tempArr

    Set Create = Me

End Function

' ━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' 配列
'
Public Property Get Elements() As Variant
Attribute Elements.VB_Description = "クラスの操作対象とする配列"
Attribute Elements.VB_UserMemId = 0

    Elements = pArray

End Property
Public Property Let Elements(ByVal aArray As Variant)
Attribute Elements.VB_Description = "クラスの操作対象とする配列"
Attribute Elements.VB_UserMemId = 0

    pArray = aArray

End Property
Public Property Set Elements(ByVal aArray As Variant)
Attribute Elements.VB_Description = "クラスの操作対象とする配列"
Attribute Elements.VB_UserMemId = 0

    Set pArray = aArray

End Property

''
' 要素
'
Public Property Get Element(ByVal Index As Long) As Variant
Attribute Element.VB_Description = "配列の要素"

    On Error GoTo ErrorCatch
    Dim e As Variant
    Select Case dimension(pArray)
        Case 1
            e = pArray(Index)

        Case Else
            GoTo ErrorCatch

    End Select

    Element = e

    Exit Property

ErrorCatch:

    MsgBox "配列の次元数を確認してください。"

End Property
Public Property Let Element(ByVal Index As Long, _
                            ByVal assginValue As Variant)
Attribute Element.VB_Description = "配列の要素"

    pArray(Index) = assginValue

End Property
Public Property Set Element(ByVal Index As Long, _
                            ByRef assginValue As Variant)
Attribute Element.VB_Description = "配列の要素"

    Set pArray(Index) = assginValue

End Property

''
' 要素数
'
Public Property Get Length() As Long
Attribute Length.VB_Description = "配列の要素数"
    Length = UBound(pArray)
End Property

''
' 配列の次元数
'
' 【Excel VBA】配列の次元数を取得する関数コード | oki2a24  Just another oki2a24 ブロゴ
' https://oki2a24.com/2016/03/09/get-dimension-number-in-excel-vba/
'
' @param {Variant} 配列
'
' @return {Long} 次元数
'
Private Property Get dimension(ByRef arr As Variant) As Long

    Dim d As Long: d = 1

    On Error Resume Next

    '次元数を確認する
    Dim tempLength As Long
    Do Until Err.Number <> 0
        tempLength = UBound(arr, d)
        d = d + 1
    Loop

    '値を調整して返す
    '（エラー発生時に次元 + 1 となっているので、2を引く）
    dimension = d - 2

End Property

' ━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' 配列を拡張する
'
' @param {Long} [maxIndex] 配列の要素数
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Extend(Optional ByVal maxIndex As Long = 0) As Array1d
Attribute Extend.VB_Description = "配列の最大要素数を増やす"

    pArray = privatelyExtend(pArray)

    Set Extend = Me

End Function

''
' 配列を拡張する
'
' @param {Variant} srcArray   拡張する配列
' @param {Long}    [maxIndex] 配列の要素数
'
' @return {Variant} 拡張後の配列
'
Private Function privatelyExtend(ByVal srcArray As Variant, _
                                 Optional ByVal maxIndex As Long = 0 _
                                 ) As Variant
    '拡張サイズを確認
    Dim extension As Long
    If maxIndex = 0 Then '判定方法…？
        extension = UBound(srcArray) + 1

    Else
        extension = maxIndex

    End If

    '値を保持したまま拡張
    ReDim Preserve srcArray(extension)

    privatelyExtend = srcArray

End Function

''
' 最後尾に要素を追加
'
' 【参照】VBA 動的配列で要素数を意識せずに拡張する方法 | t-hom’s diary
' https://thom.hateblo.jp/entry/2015/03/19/213019
'
' @param {Variant} 代入値
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Push(ByVal var As Variant) As Array1d
Attribute Push.VB_Description = "配列の末尾に値を挿入する"

    '配列を拡張
    Me.Extend

    '最後尾に追加
    pArray = privatelyPush(pArray, var)

    Set Push = Me

End Function

''
' 最後尾に要素を追加
'
' 【参照】VBA 動的配列で要素数を意識せずに拡張する方法 | t-hom’s diary
' https://thom.hateblo.jp/entry/2015/03/19/213019
'
' @param {Variant} srcArray 編集する配列
' @param {Long}    var      追加する要素
'
' @return {Variant} 追加後の配列
'
Private Function privatelyPush(ByRef srcArray As Variant, _
                               ByVal var As Variant) As Variant

'    '配列を拡張
'    srcArray = privatelyExtend(srcArray)

    '最後尾に追加
    srcArray(UBound(srcArray)) = var

    privatelyPush = srcArray

End Function

''
' 配列を反転する
'
' 【参考】
' VBA 配列の並びを反転させる
' https://www.tipsfound.com/vba/02021
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Reverse() As Array1d
Attribute Push.VB_Description = "配列を反転する"

    Dim low As Long:   low = LBound(pArray)
    Dim high As Long: high = UBound(pArray)
    Dim arrLength As Long: arrLength = Me.Length

    '反転した配列を作成
    Dim reversedArray() As Variant
    ReDim Preserve reversedArray(low To high)
    Dim i As Long
    For i = 0 To arrLength
        reversedArray(low + i) = pArray(high - i)
    Next

    '元の配列に代入
    For i = 0 To arrLength
        pArray(low + i) = reversedArray(low + i)
    Next

    Set Reverse = Me

End Function

''
' 配列内の要素をさらに区切る
'
' @param {Long} [maxIndex] インデックスの最大値
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Subdivide(Optional ByVal delimiter As String = ", ") As Array1d
Attribute Push.VB_Description = "特定の文字列で要素を区切る"

    Dim newArray As Variant: ReDim newArray(0)
    Dim tmpArray As Variant: ReDim tmpArray(0)
    Dim newIndex As Long: newIndex = 0

    Dim i As Long
    For i = 0 To UBound(pArray)
        tmpArray = Split(pArray(i), delimiter)

        Dim t As Long
        For t = 0 To UBound(tmpArray)
            newArray = privatelyAdd(newArray, tmpArray(t), newIndex)

            newIndex = newIndex + 1
        Next t
    Next i

    pArray = newArray

End Function

''
' 要素を追加する（Extendメソッド、Pushメソッドの代用）
'
' @param {Variant} srcArray            元の配列
' @param {Variant} var                 挿入する値
' @param {Long}    [insertionPosition] 挿入する位置
'
' @return {Variant} 追加後の配列
'
Private Function privatelyAdd(ByVal srcArray As Variant, _
                              ByVal var As Variant, _
                              Optional ByVal insertionPosition As Long = 0 _
                              ) As Variant

    '配列を拡張
    If insertionPosition <> 0 Then
        srcArray = privatelyExtend(srcArray, insertionPosition)
    End If

    '最後尾に追加
    srcArray = privatelyPush(srcArray, var)

    privatelyAdd = srcArray

End Function

''
' 全要素の値から特定の文字列を除去する
'
' @note
'   今のところ、要素が「すべてString型」であることを前提としている。
'
' @param {String} erasing 除去する文字列
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Eraser(ByVal erasing As String) As Array1d
Attribute Eraser.VB_Description = "全要素から特定の文字列を除去する"

    Dim i As Long
    For i = 0 To UBound(pArray)
        pArray(i) = Replace(pArray(i), erasing, "")
    Next i

End Function

''
' 配列を連結する
'
' @param {String} [delimiter] 区切り文字
'
' @return {String} 連結した文字列
'
Public Function JoinVia(Optional ByVal delimiter As String = ", ") As String
Attribute Eraser.VB_Description = "配列の要素を1連の文字列にする"

    Dim str As String: str = Join(pArray, delimiter)
    If str = "" Then str = "EMPTY"

    JoinVia = str

End Function

''
' 値をログ出力する
'
' @param {String} [delimiter] 区切り文字
'
Public Sub PrintAll(Optional ByVal delimiter As String = ", ")
Attribute Eraser.VB_Description = "配列の要素をログ出力する"

    Dim str As String: str = JoinVia(delimiter)

    Debug.Print "[" & str & "]"

End Sub
