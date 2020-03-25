VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'━━━━━━━━ プロパティ ━━━━━━━━━

Public Property Get Apps() As Application:             Set Apps = pApps: End Property
Public Property Set Apps(ByVal aApps As Application): Set pApps = aApps: End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'終了時処理
Private Sub Class_Terminate()
'    使うときは先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'━━━━━━━━ メソッド ━━━━━━━━━

''
' ステータスバーに進捗を表示する
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
Public Sub ShowProgressWithPercentage(ByVal currentCounts As Long, _
                                      ByVal totalCounts As Long, _
                                      Optional ByVal description As String = "出力")
    
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
        
        Apps.StatusBar = m
    End With

End Sub

''
' ステータスバーに進捗を表示する
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
Public Sub ShowProgressWithNumbers(ByVal currentCounts As Long, _
                                   ByVal totalCounts As Long, _
                                   Optional ByVal units As String = "件", _
                                   Optional ByVal description As String = "出力")
    
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
        
        Apps.StatusBar = m
    End With

End Sub

''
' ステータスバーにアニメーションを表示する
'
' @param {Long} currentCounts 現在の件数
' @param {Long} divisor       除数
'
Public Sub ShowWaitingAnime(ByVal currentCounts As Long, _
                            Optional ByVal msg As String = "データを解析しています", _
                            Optional ByVal divisor As Long = 1000)

    With Me
        'アニメーション風
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

        With pApps
            .StatusBar = msg & dotMark
        End With
    End With

End Sub
