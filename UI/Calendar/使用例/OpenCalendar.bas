Attribute VB_Name = "OpenCalendar"
Option Explicit

'カレンダー用の変数
Public clickedDate As Date      'テキストボックスの値を格納する変数
Public isDateClicked As Boolean 'カレンダーがクリックされたか判定するフラグ

Sub OpenForm()
  
  With MainForm
    .TextBox1 = Format(Date, "yyyy/mm/dd") '初期値
    .Show
  End With

End Sub
