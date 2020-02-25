VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'━━━━━━━━ モジュールレベル変数 ━━━━━━━━━

Private pApps As Application

'━━━━━━━━ プロパティ ━━━━━━━━━

Property Get Apps() As Application: Set Apps = pApps: End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'━━━━━━━━ メソッド ━━━━━━━━━

Public Sub Start()

    '画面の再描画/自動計算/イベント受付を停止
    With Me.Apps
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
    End With

End Sub
