VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "1行分をデータのコレクションとして扱います。"
Option Explicit

' ━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pRecord As Collection

' ━━━━━━━━━━━ コンストラクタ等 ━━━━━━━━━━━

''
' 初期化
'
Private Sub Class_Initialize()

    Set pRecord = New Collection

End Sub

''
' 1行分の値をコレクションとして取り込む
'
' @param {Range} Row 1行分の範囲
'
' @return {Array2d} Me（メソッドチェーンを可能にするため）
'
Public Function Import(ByRef Row As Range) As Record

    'メモリ節約のため使用範囲だけ取得する
    Dim importRange As Range
    With Row
        Set importRange = .Range( _
                              .Columns(1), _
                              .Cells(1, .Columns.Count).End(xlToLeft) _
                          )
    End With

    '1セルずつ格納
    Dim i As Long
    For i = 1 To importRange.Columns.Count
        pRecord.Add importRange.Columns(i).Value
    Next i

    Set Import = Me

End Function

' ━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

Public Property Get FieldCounts() As Long

    FieldCounts = pRecord.Count

End Property

''
' レコードを取得する
'
' @return {Collection} レコード1行分
'
Property Get Items() As Collection

    Set Items = pRecord

End Property

' ━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' レコードから1セル分の値を取得する
'
' @param {Long} Itemのキー（数値）
'
' @return {Variant} レコード1行分
'
Public Function Item(ByVal key As Long) As Variant
    Attribute Item.VB_Description = "レコードから1セル分の値を取得する"
    Attribute Item.VB_UserMemId   = 0

    Item = pRecord.Item(key)

End Function

''
' 要素を追加
'
' @param {Variant} Value 追加する値
'
' @return {Variant} レコード1行分
'
Public Function Add(ByRef Value As Variant, _
                    Optional ByVal fieldName As String) As Record

    If IsMissing(fieldName) Then
        pRecord.Add Value

    Else
        pRecord.Add Value, fieldName

    End If

    Set Add = Me

End Function

''
' レコードを反転する
'
' 【参考】
' VBA 配列の並びを反転させる
' https://www.tipsfound.com/vba/02021
'
' @return {Record} Me（メソッドチェーンを可能にするため）
'
Public Function Reverse() As Record

    Dim fields As Long: fields = pRecord.Count

    '反転したコレクションを作成
    Dim reversed As Collection: Set reversed = New Collection
    Dim i As Long
    For i = fields To 1 Step -1
        reversed.Add pRecord(i)
    Next i

    '一旦初期化
    '（Itemメソッドで代入はできないため）
    Set pRecord = New Collection

    '元のコレクションに戻す
    For i = 1 To fields
        pRecord.Add reversed(i)
    Next i

    Set Reverse = Me

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

''
' 配列を連結する
'
' @param {string} 区切り文字
'
' @return {string} 連結した文字列
'
Public Function JoinVia(Optional ByVal delimiter As String = ", ") As String

    Dim str As String: str = Join(ToArray(), delimiter)
    If str = "" Then str = "EMPTY"

    JoinVia = str

End Function

''
' コレクションの値を一次元配列に変換する
'
' 【参考】
' How to join a collection in VBA - Stack Overflow
' https://stackoverflow.com/questions/29015444/how-to-join-a-collection-in-vba
'
' @return {Varinat} 生成した一次元配列
'
Public Function ToArray() As Variant

    Dim fields As Long: fields = pRecord.Count

    Dim arr As Variant: ReDim arr(fields - 1)
    Dim i As Long
    For i = 1 To fields
        arr(i - 1) = pRecord(i)
    Next i

    ToArray = arr

End Function
