VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "ステータスバーを制御してマクロの進捗等を表示します。"
Option Explicit

'━━━━━━━━ モジュールレベル変数 ━━━━━━━━━

'列挙体
'Public Enum e列挙体
'    メンバー = 1
'End Enum

'構造体
'Public Type t構造体
'    メンバー As 型
'End Type

'モジュールレベル変数
Private pApps As Application
Private pMessage As String

'━━━━━━━━ プロパティ ━━━━━━━━━

Public Property Get Apps() As Application
Attribute Apps.VB_Description = "ステータスバーで使用するApplicationオブジェクト"

    Set Apps = pApps

End Property
Public Property Set Apps(ByRef aApps As Application)
Attribute Apps.VB_Description = "ステータスバーで使用するApplicationオブジェクト"

    Set pApps = aApps

End Property

''
' デフォルトで表示する文字列
'
Public Property Get Message() As String
Attribute Message.VB_Description = "デフォルトで表示する文字列"
Attribute Message.VB_UserMemId = 0

    Message = pMessage

End Property
Public Property Let Message(ByVal msg As String)
Attribute Message.VB_Description = "デフォルトで表示する文字列"
Attribute Message.VB_UserMemId = 0

    pMessage = msg

End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

''
' 初期化処理
'
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "クラスの初期化"

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

''
' 終了時処理
'
Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "クラスの後始末"

'    使うときは先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."

End Sub

'━━━━━━━━ メソッド ━━━━━━━━━

''
' ステータスバーにメッセージを点滅させる
'
' @param {String} msg 表示する文字列
'
Public Sub Flash(Optional ByVal msg As String, _
                 Optional ByVal canResumeScreen As Boolean = False)
Attribute Eraser.VB_Description = "ステータスバーにメッセージを点滅させる"

    If msg = "" Then msg = pMessage

    With pApps
        .ScreenUpdating = True

        .StatusBar = False
        .StatusBar = msg

        .ScreenUpdating = canResumeScreen
    End With

End Sub

''
' ステータスバーにメッセージを表示する
'
' @param {String} msg 表示する文字列
'
Public Sub Display(Optional ByVal msg As String, _
                   Optional ByVal canResumeScreen As Boolean = False)
Attribute Eraser.VB_Description = "ステータスバーにメッセージを表示する"

    If msg = "" Then msg = pMessage

    With pApps
        .ScreenUpdating = True

        .StatusBar = msg

        .ScreenUpdating = canResumeScreen
    End With

End Sub

''
' ステータスバーの表示を解除する
'
Public Sub Clear()
Attribute Clear.VB_Description = "ステータスバーの表示を解除する"

    pApps.StatusBar = False

End Sub

''
' 進捗をパーセンテージで表現する
'
' 【参考】
' Excel VBA 時間のかかるマクロの進捗状況を表示する | MyRecord
' https://kosapi.com/post-2305/
'
' @param {Long}   currentCounts 現在の件数
' @param {Long}   totalCounts   処理予定件数
' @param {String} units         単位表記
' @param {String} description   処理内容
'
' @return {String} 現在の進捗
'
Public Property Get GraphWithPercentage(ByVal currentCounts As Long, _
                                        ByVal totalCounts As Long, _
                                        Optional ByVal description As String _
                                                       = "出力" _
                                        ) As String
Attribute GraphWithPercentage.VB_Description = "進捗をパーセンテージで表現する"

    With Me
        'グラフ表示用
        Dim currentSymbols As String
        Const MAX_COUNTS As Long = 10
        Dim blackCounts As Long
        blackCounts = Int((currentCounts / totalCounts) * MAX_COUNTS)
        currentSymbols = String(blackCounts, "■") & String(MAX_COUNTS, "□")

        'ステータス文
        Dim m As String
        m = description & "中..."
        m = m & " " & Left(currentSymbols, MAX_COUNTS)
        m = m & " " & Format(currentCounts / totalCounts, "0 %")
    End With

    GraphWithPercentage = m

End Property

''
' 進捗を件数で表現する
'
' 【参考】
' Excel VBA 時間のかかるマクロの進捗状況を表示する | MyRecord
' https://kosapi.com/post-2305/
'
' @param {Long}   currentCounts 現在の件数
' @param {Long}   totalCounts   処理予定件数
' @param {String} units         単位表記
' @param {String} description   処理内容
'
' @return {String} 現在の進捗
'
Public Property Get GraphWithCounts(ByVal currentCounts As Long, _
                                    ByVal totalCounts As Long, _
                                    Optional ByVal units As String = "件", _
                                    Optional ByVal description As String _
                                                   = "出力" _
                                    ) As String
Attribute GraphWithCounts.VB_Description = "進捗を件数で表現する"

    With Me
        'グラフ表示用
        Dim currentSymbols As String
        Const MAX_COUNTS As Long = 10
        Dim blackCounts As Long
        blackCounts = Int((currentCounts / totalCounts) * MAX_COUNTS)
        currentSymbols = String(blackCounts, "■") & String(MAX_COUNTS, "□")

        'ステータス文
        Dim m As String
        m = description & "中..."
        m = m & " " & Left(currentSymbols, MAX_COUNTS)
        m = m & " " & Format(currentCounts, "#,##0") & " " & units
        m = m & " / " & Format(totalCounts, "#,##0") & " " & units
    End With

    GraphWithCounts = m

End Property

''
' ステータスバーに表示するアニメーション
'
' @param {Long} currentCounts 現在の件数
' @param {Long} divisor       除数
'
''
' ステータスバーにアニメーションを表示する
'
' @param {Long} currentCounts 現在の件数
' @param {Long} divisor       除数
'
Public Property Get Animation(ByVal currentCounts As Long, _
                              Optional ByVal msg As String _
                                             = "データを解析しています", _
                              Optional ByVal divisor As Long = 1000)
Attribute Animation.VB_Description = "ステータスバーに表示するアニメーション"

    'ドットの増減でアニメーションを表現する
    With Me
        Dim dotMark As String
        Dim criteria As Double: criteria = currentCounts / divisor
        Select Case criteria
            Case Is < 3
                dotMark = "  "
            Case Is < 6
                dotMark = ".  "
            Case Is < 9
                dotMark = ".. "
            Case Is >= 9
                dotMark = "..."
        End Select
    End With

    Animation = msg & dotMark

End Property
