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
' コンストラクタ
'
' @param {Long} maxIndex インデックスの最大値
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Init(Optional ByVal maxIndex As Long = 0) As Array1d

    ReDim pArray(maxIndex) '配列の初期化

    Set Init = Me

End Function

' ━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' 配列
'
Public Property Get arr() As Variant
    arr = pArray
End Property
Public Property Let arr(ByVal aArray As Variant)
    pArray = aArray
End Property
Public Property Set arr(ByVal aArray As Variant)
    Set pArray = aArray
End Property

'''
'' 多次元配列を一次元配列に変換する
''
'' @param {●型●} name ●概要●
'' @param {●型●} name ●概要●
'' @param {●型●} name ●概要●
''
'' @return {●型●} ●概要●
''
'Public Function To1d() As Variant
'
'    Dim maxIndex As Long: maxIndex = UBound(pArray, dimension(pArray))
'
'    Dim arr1d As Variant
'    ReDim arr1d(maxIndex) '配列の初期化
'
'
'
'End Function

''
' 配列の次元数をログ出力する
'
' 【Excel VBA】配列の次元数を取得する関数コード | oki2a24  Just another oki2a24 ブロゴ
' https://oki2a24.com/2016/03/09/get-dimension-number-in-excel-vba/
'
' @param {variant} 配列
'
' @return {long} 次元数
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

''
' 要素
'
Public Property Get Elements(ByVal Index As Long) As Variant

    On Error GoTo ErrorCatch
    Dim e As Variant
    Select Case dimension(pArray)
        Case 1
            e = pArray(Index)

        Case 2
            e = pArray(1, Index)

        Case Else
            GoTo ErrorCatch

    End Select

    Elements = e

    Exit Property

ErrorCatch:

    MsgBox "配列の次元数が適切でないかもしれません。"

End Property
Public Property Let Elements(ByVal Index As Long, _
                             ByVal assginValue As Variant)
    pArray(Index) = assginValue
End Property
Public Property Set Elements(ByVal Index As Long, _
                             ByRef assginValue As Variant)
    Set pArray(Index) = assginValue
End Property

''
' 要素数
'
Public Property Get Length() As Long
    Length = UBound(pArray)
End Property

' ━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' 配列を拡張する
'
' @param {Long} maxIndex 配列の要素数
'
' @return {Array1d} Me（メソッドチェーンを可能にするため）
'
Public Function Extend(Optional ByVal maxIndex As Long = 0) As Array1d

    '拡張サイズを確認
    Dim extension As Long
    If maxIndex = 0 Then '判定方法…？
        extension = UBound(pArray) + 1

    Else
        extension = maxIndex

    End If

    '値を保持したまま拡張
    ReDim Preserve pArray(extension)

    Set Extend = Me

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

    '配列を拡張
    Me.Extend

    '最後尾に追加
    pArray(UBound(pArray)) = var

    Set Push = Me

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
' 配列を連結する
'
' @param {string} 区切り文字
'
' @return {string} 連結した文字列
'
Public Function JoinVia(Optional ByVal delimiter As String = ", ") As String

    Dim str As String: str = Join(pArray, delimiter)
    If str = "" Then str = "EMPTY"

    JoinVia = str

End Function

''
' 値をログ出力する
'
' @param {string} 区切り文字
'
Public Sub PrintAll(Optional ByVal delimiter As String = ", ")

    Dim str As String: str = JoinVia(delimiter)

    Debug.Print "[" & str & "]"

End Sub
