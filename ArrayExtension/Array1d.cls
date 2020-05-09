VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Array1d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "配列を扱いやすくします。"
Option Explicit

' ━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pArray As Variant
Private pDelimiter As String

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

    pDelimiter = ","

    Set Init = Me

End Function

' ━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' 配列
'
Public Property Get Arr() As Variant
    Arr = pArray
End Property
Public Property Set Arr(ByVal aArray As Variant)
    pArray = aArray
       Arr = pArray
End Property

''
' 要素
'
Public Property Get Elements(ByVal Index As Long) As Variant
    Elements = pArray(Index)
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
     If str = "" Then: str = "EMPTY"

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
