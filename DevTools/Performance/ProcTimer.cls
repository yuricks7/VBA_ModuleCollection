VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'━━━━━━━━━ モジュール変数 ━━━━━━━━━━

Private pApps As Application
Private pStartDate As Single

Private pTime As Double
Private pStartTime As Double
Private pFinishTime As Double

'━━━━━━━━━ プロパティ ━━━━━━━━━━

Property Get Apps() As Application: Set Apps = pApps: End Property
Property Get StartDate() As Double: StartDate = pStartDate: End Property

Property Get Time() As Double
    With Me
        pTime = .FinishTime - .StartTime
    End With
    Time = pTime
End Property

Property Get StartTime() As Double: StartTime = pStartTime: End Property

Property Get FinishTime() As Double: FinishTime = pFinishTime: End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()

    Set pApps = Application
    pStartDate = Now

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'終了時処理
Private Sub Class_Terminate()

    ' このあと変数を破棄してしまうので、先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'━━━━━━━━━━ メソッド ━━━━━━━━━━

''
' スタート
'
Public Sub Start()
    pStartTime = Timer
End Sub

'''
'' 一時停止
''
'Public Sub Pause()
'
'    '作成中
'
'End Sub

''
' カウントをスタート
'
Public Sub Finish()
    pFinishTime = Timer
End Sub

''
' 処理時間を記録
'
' @param {Worksheet} targetSheet   記録先
' @param {Long}      trialTimes    試行回数
' @param {String}    procedureName プロシージャ名
'
Public Sub Log(ByRef targetSheet As Worksheet, _
               Optional ByVal trialTimes As Long = 1, _
               Optional ByVal procedureName As String = "")
    
    If targetSheet Is Nothing Then
        MsgBox "引数にシートが設定されていません。"
        Stop
    End If
    
    With Me
        If .FinishTime = 0 Then .Finish
    
        Dim values() As Variant
        values = Array(.StartDate, .Time, trialTimes, procedureName)
        End With
        
        With targetSheet
            Dim lastRow As Long
            
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            .Cells(lastRow, 1).Resize(1, UBound(values) + 1) = values
'            .Range(.Cells(lastRow, 1), .Cells(lastRow, UBound(values) + 1)).Value = values
    End With

End Sub
