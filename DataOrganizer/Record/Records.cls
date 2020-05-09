VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pRecord As Record
Private pRecords As Collection

Private pValues As Variant

' ━━━━━━━━━━━ コンストラクタ等 ━━━━━━━━━━━

''
' 初期化
'
Private Sub Class_Initialize()
    
    Set pRecords = New Collection

End Sub

''
' シートの値を取り込む
'
' @param {Worksheet} srcSheet ワークシート
'
' @return {Records} Me（メソッドチェーンを可能にするため）
'
Public Function Import(ByRef srcSheet As Worksheet) As Records

    Me.Values = srcSheet.UsedRange.Value
    Call importRowsData(srcSheet)
    
    Set Import = Me
    
End Function

''
' シートの値をArray1dクラスのコレクションに変換する
'
' @param {Worksheet} srcSheet ワークシート
'
Private Sub importRowsData(ByRef srcSheet As Worksheet)

    'メモリ節約のため、使用済み範囲のみとする
    Dim importRange As Range
    With srcSheet.Cells
        Set importRange = .Range( _
                              .Rows(1), _
                              .Cells(.Rows.Count, 1).End(xlUp) _
                          )
    End With

    Dim i As Long
    For i = 1 To importRange.Rows.Count
        Dim r As Record: Set r = New Record
        r.Import (srcSheet.Rows(i))
        
        pRecords.Add r
    Next i
    
End Sub

' ━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' シートの値（二次元配列）
'
Public Property Get Values() As Variant
    Values = pValues
End Property
Public Property Let Values(ByVal aValues As Variant)
    pValues = aValues
End Property

''
' レコードを取得する
'
' @return {Collection} シート1枚分
'
Property Get Items() As Collection

    Set Items = pRecords

End Property

' ━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' レコードから1セル分の値を取得する
'
' @param {Long} Itemのキー（数値）
'
' @return {Variant} レコード1行分
'
Public Function Item(ByVal key As Long) As Record

    Set Item = pRecords.Item(key)

End Function
