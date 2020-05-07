VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "PowerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "PowerQueryを使ってデータを更新します。"

Option Explicit

'━━━━━━━━ 値の定義 ━━━━━━━━━

'列挙体
'Public Enum e列挙体
'
'End Enum

'構造体
'Public Enum t構造体
'
'End Enum

'━━━━━━━━ モジュールレベル変数 ━━━━━━━━━

Private pQueryTable As QueryTable
Private pQueryTables As QueryTables

Private pTargetBook As Workbook

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()


End Sub

''
' コンストラクタ
'
' @param {workbook} aTargetBook クエリを実行するブック
'
' @return {PowerQuery} Me
'
Public Function Init(ByRef aTargetBook As Workbook) As PowerQuery

    Set pTargetBook = aTargetBook
    
    Set Init = Me

End Function

'終了時処理
Private Sub Class_Terminate()
'    使うときは先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'━━━━━━━━ プロパティ ━━━━━━━━━

Property Get TargetBook() As Workbook: Set TargetBook = pTargetBook: End Property

''
' ブック内の「バックグラウンドクエリ」の可否をまとめて切り替える
'
' 【注意】
' "True" にすると、クエリはマクロの終了後に実行される。
' （実行中はファイルに触れないので、どうせなら "False" にしてマクロの中で実行した方が良いかも。）
'
' @param {boolean} isTrue
'
Public Property Let CanRunBackgroundQueries(ByVal isTrue As Boolean)

    With Me.TargetBook
        Dim i As Long
        For i = 1 To .Connections.Count
          'Excludes PowerPivot and other connections
          If .Connections(i).Type = xlConnectionTypeOLEDB Then
            .Connections(i).OLEDBConnection.BackgroundQuery = isTrue
          End If
        Next i
    End With

End Property

'━━━━━━━━ メソッド ━━━━━━━━━

''
' 「すべて更新」の実行
'
Public Sub RefreshAll(Optional ByRef targetSheet As Worksheet)
    'すべて更新
    Me.TargetBook.RefreshAll
    
    'シートの後始末
    If Not targetSheet Is Nothing Then Call adjustRowHeights(targetSheet)

End Sub

''
' PowerQueryで設定したクエリのみ実行する
'
' @param {String}    queryName 内容
' @param {Worksheet} targetSheet 内容
'
Public Sub ImportData(ByVal queryName As String, _
                      Optional ByRef targetSheet As Worksheet)

    'データの更新
    On Error GoTo Catch:

    With pTargetBook
        .Connections(queryName).Refresh
    End With
    
    'シートの後始末
    If Not targetSheet Is Nothing Then Call adjustRowHeights(targetSheet)

    Exit Sub

Catch:

    MsgBox "データのインポート中に何らかの問題がありました。" & vbCrLf _
         & "処理を中止します。", vbInformation

End Sub

''
' シートの見栄えを調整する
'
' @param {Worksheet} targetSheet 対象のワークシート
'
Private Sub adjustRowHeights(ByRef targetSheet As Worksheet)

    With targetSheet
        With .Rows
            .Font.Size = 8
            .RowHeight = 14.4 '24px
            .VerticalAlignment = xlBottom '下揃え
        End With
        
        .Rows(1).Font.Size = 9 'ヘッダー
    End With

End Sub
