VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "PerformanceBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = true
Attribute VB_Exposed         = False
Attribute VB_Description     = "Excelマクロ高速化にあたり、実行速度に影響するものを制御します。"
Option Explicit

'━━━━━━━━ モジュールレベル変数 ━━━━━━━━━

Private pApps As Application

'━━━━━━━━ プロパティ ━━━━━━━━━

Property Get Apps() As Application
Attribute Apps.VB_Description _= "このクラスで使用するApplicationオブジェクト"

    Set Apps = pApps

End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "クラスの初期化"

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'━━━━━━━━ メソッド ━━━━━━━━━

Public Sub Start()
Attribute Start.VB_Description = "高速化スタート"
Attribute Start.VB_UserMemId   = 0

    '画面の再描画/自動計算/イベント受付を停止
    With Me.Apps
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

End Sub

Public Sub Finish()
Attribute Finish.VB_Description = "後始末"

    '画面の再描画/自動計算/イベント受付を再開
    With Me.Apps
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub
