VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArray() As Variant
Private pLength As Long
Private pDimension As Long

Property Get arr() As Variant
    arr = pArray
End Property

Property Set arr(ByRef aArray As Variant)
    pArray = aArray
End Property

Property Get Length() As Long
    Length = UBound(pArray)
End Property

''
' コンストラクタ
'
' @param {variant} 処理対象の配列
'
' @return {object} Me
'
Function Init(Optional ByVal uBoundIndex As Long = 0) As PowerArray

    ReDim pArray(uBoundIndex) '配列の初期化

    Set Init = Me
End Function

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
Property Get Dimension() As Long
    Dim temp As Long
    temp = 1

    Dim forDecision As Long

    '次元数を確認
    On Error Resume Next
    Do While Err.Number = 0 'エラーが起きたらストップ
        forDecision = UBound(arr, temp)
        temp = temp + 1
    Loop

    '値の調整
    '（エラー発生時に次元 + 1 となっているので、2を引く）
    pDimension = temp - 2

    Dimension = pDimension
End Property

''
' 最後尾に要素を追加
'
' 【参照】VBA 動的配列で要素数を意識せずに拡張する方法 | t-hom’s diary
' https://thom.hateblo.jp/entry/2015/03/19/213019
'
' @param {variant} 追加する値
'
' @return {variant} 追加後の配列
'
Function Push(ByVal var As Variant) As Variant

    '最後尾に追加
    pArray(UBound(pArray)) = var

    '配列を拡張
    ReDim Preserve pArray(UBound(pArray) + 1)

    Push = pArray

End Function

''
' 指定のインデックス位置に値を代入
'
' @param {long} 対象のインデックス
' @param {variant} 追加する値
'
' @return {variant} 追加後の配列
'
Function Assign(ByVal targetIndex As Long, _
                ByVal var As Variant) As Variant

    pArray(targetIndex) = var

    Assign = pArray

End Function

''
' 配列を反転する
'
' @param {string} 区切り文字
'
' @return {string} 連結した文字列
'
Function Reverse() As Variant

    Dim low As Long: low = LBound(pArray)
    Dim high As Long: high = UBound(pArray)

    Dim tempArray() As Variant
    ReDim Preserve tempArray(low To high)

    Dim arrLength As Long
    arrLength = high - low

    Dim i As Long
    For i = 0 To arrLength
        tempArray(low + i) = pArray(high - i)
    Next

    For i = 1 To arrLength
        pArray(low + i) = tempArray(low + i)
    Next

    Reverse = pArray
End Function

''
' 2層になった配列を二次元配列に置き換える
' 【参照】
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @param {array} [arr(0)(0)]式の配列
'
' @return {array} [arr(0,0)]式の配列
'
Private Function get2dValues(ByRef nest2dArr As Variant) As Variant

    Dim ret() As Variant
    ReDim ret(0 To UBound(nest2dArr), 0 To UBound(nest2dArr(0)))

    Dim r As Long: r = 0
    Dim rowData As Variant
    For Each rowData In nest2dArr
        Dim c As Long: c = 0
        Dim element As Variant
        For Each element In rowData
            ret(r, c) = element
            c = c + 1
        Next
        r = r + 1
    Next

    get2dValues = ret
End Function

''
' 配列を連結する
'
' @param {string} 区切り文字
'
' @return {string} 連結した文字列
'
Function JoinVia(ByVal delimiter As String) As String

    Dim str As String
    str = Join(pArray, delimiter)

    If str = "" Then
        JoinVia = "EMPTY"
        Exit Function
    End If

    Dim Length As Long
    Length = Len(str)

    '最後の区切り文字を削除
    str = Left(str, Length - Len(delimiter))
    JoinVia = str

End Function

''
' 改行を付けながら値をすべて連結する
'
' @return {string} 連結した文字列
'
Function JoinWithCRLF() As String
    '初期化
    Dim str As String: str = ""

    '2つ目～最後尾1つ前まで、改行と結合
    Dim a As Long
    For a = 1 To UBound(arr) - 1
        '1つ目
        If a = 0 Then
            str = pArray(0)

        '最後は値の追加のみ
        ElseIf a = UBound(pArray) Then
            str = str & pArray(UBound(pArray))

        'その他は、値+改行を結合
        Else
            str = str & vbCrLf & CStr(pArray(a))
        End If
    Next

    JoinWithCRLF = str
End Function

''
' 値をログ出力する
'
' @param {string} 区切り文字
'
Sub PrintAll(ByVal delimiter As String)

    Dim p As String
    p = JoinVia(delimiter)

    Debug.Print "[" & p & "]"

End Sub
