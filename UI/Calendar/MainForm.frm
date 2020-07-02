VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "日付選択"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3675
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
  
  isDateClicked = False 'リセット
  
  '日付の有無で処理を分ける
  If IsDate(Me.TextBox1) Then
    clickedDate = Me.TextBox1 'テキストボックスの日付を格納
  
  Else
    clickedDate = Date '今日の日付を格納
  
  End If
  
  'カレンダーでクリックされた日付を入力
  CalendarForm.Show
  If isDateClicked Then Me.TextBox1 = Format(clickedDate, "yyyy/mm/dd")

End Sub
