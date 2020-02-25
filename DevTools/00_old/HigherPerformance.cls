VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HigherPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'━━━━━━━━━ プロパティ ━━━━━━━━━━

Private pApps As Application
Private pStartDate As Double

Private pTime As Double
Private pStartTime As Double
Private pFinishTime As Double

Property Get Apps() As Application: Set Apps = pApps: End Property
Property Get StartDate() As Double
    StartDate = pStartDate
End Property

Property Get Time() As Double
    With Me
        pTime = .FinishTime - .StartTime
    End With
    Time = pTime
End Property

Property Get StartTime() As Double:   StartTime = pStartTime:  End Property
Property Get FinishTime() As Double: FinishTime = pFinishTime: End Property


'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()

    Set pApps = Application
    pStartDate = Now()

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'終了時処理
Private Sub Class_Terminate()

    ' このあと変数を破棄してしまうので、先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."

End Sub

'━━━━━━━━━━ メソッド ━━━━━━━━━━
Public Sub Start()

    '画面の再描画/自動計算/イベント受付を停止
    With Me.Apps
        .StatusBar = "作業開始しますー(･0･)ゞ"
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

End Sub

Public Sub Finish()

    '画面の再描画/自動計算/イベント受付を再開
    With Me.Apps
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .Wait [Now()] + 50 / 86400000 '描画再開に時間がかかることがあるため
        .StatusBar = "おまたせしましたー(･0･)ゞ"
    End With

End Sub

Public Sub SetTimer()
    pStartTime = Timer
End Sub

Public Sub PauseTimer()



End Sub

Public Sub StopTimer()
    pFinishTime = Timer
End Sub


Public Sub Log(Optional ByRef targetSheet As Worksheet, _
               Optional ByVal trialTimes As Long = 1, _
               Optional ByVal procedureName As String = "")
    
'    If targetSheet Is Nothing Then
'        Set targetSheet = shProcLog
'    End If

    With Me
        If .FinishTime = 0 Then .StopTimer
    
        Dim values() As Variant
        values = Array(.StartDate, .Time, trialTimes, procedureName)
        End With
        
        With targetSheet
            Dim lastRow As Long
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            .Range(.Cells(lastRow, 1), .Cells(lastRow, UBound(values) + 1)).Value = values
    End With
End Sub


Public Sub CursorMovesTo(ByRef startPosition As Range)

    '▼前処理
    '画面描画が止まっている場合、カーソルだけ移動して
    'ウィンドウ位置はそのままになってしまうので、
    '念のため画面描画を再開しておく
    Application.ScreenUpdating = True

    '「セル.Select」のエラー防止
    startPosition.Parent.Activate

    '▼メイン
    With startPosition
        '（独立させた方がいい…？）
        If .Value = "" Then Set startPosition = .End(xlDown)
    End With

    '再代入が実行されても当初からWithブロックで選択されているstartPositionは
    '再代入前の値のままなので、Withの外側で改めて選択
    startPosition.Select

    '▼後処理
    '元に戻しておく
    Application.ScreenUpdating = False

End Sub


Public Sub ProgressBar(ByVal totalCounts As Long, _
                       ByVal currentCounts As Long)
    With Me
        '現在の時間を計算
        .StopTimer
        Dim remained As Double
        remained = .Time / currentCounts * (totalCounts - currentCounts)

        Apps.StatusBar = "出力中... [済] " & Format(currentCounts, "#,##0") & " 件 / " & Format(totalCounts, "#,##0") & " 件"
    End With

End Sub


Public Sub WaitingAnime(Optional ByVal currentCounts As Long = 300)

    With Me
        'アニメ
        Dim dotMark As String
        Dim quotient As Double
        quotient = currentCounts / 3000
        Select Case quotient
            Case Is < 3
                dotMark = ""
            Case Is < 6
                dotMark = "."
            Case Is < 9
                dotMark = ".."
            Case Is >= 9
                dotMark = "..."
        End Select

        With .Apps
            .StatusBar = "データを解析しています" & dotMark
        End With
    End With

End Sub
