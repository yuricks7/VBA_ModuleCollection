VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' イベント処理用の変数
'
Private WithEvents Target As MSForms.Label
Attribute Target.VB_VarHelpID = -1

''
' ボタンを格納
'
Public Sub Add(newControl As MSForms.Label)
    Set Target = newControl
End Sub

''
' クリックイベント
'
Private Sub Target_Click()
    
    If Target.Caption = "" Then Exit Sub 'ラベルが空だったら中止
    
    '日付を生成して変数に格納
    With CalendarForm
        clickedDate = .YearBox & "/" & .MonthBox & "/" & Target.Caption
    End With
    
    isDateClicked = True 'フラグを立てる
    
    'カレンダーを閉じて終了
    Unload CalendarForm

End Sub
