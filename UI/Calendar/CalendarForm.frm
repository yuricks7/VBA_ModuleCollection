VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "カレンダー"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' カレンダーコントロール
'
' 【参考】
' ExcelVBAでカレンダーコントロールを自作する | *Ateitexe
' https://ateitexe.com/excel-vba-calendar-control/
'------------------------------------------------------------------------------

Private Const LAST_BUTTON As Long = 42

'イベント処理用クラス
Private dayButtons(1 To LAST_BUTTON) As New DayButtonHandler

''
' 初期化
'
Private Sub UserForm_Initialize()
  
    With Me
        '年月の選択肢を登録
        Dim y As Long
        For y = -3 To 3 '前後3年分
            .YearBox.AddItem CStr((Year(clickedDate)) + y)
        Next y
        
        Dim m As Long
        For m = 1 To 12
            .MonthBox.AddItem CStr(m)
        Next m
        
        '年月の初期値を指定
        .YearBox = Year(clickedDate)
        .MonthBox = Month(clickedDate)
    
        'ラベルのクリックイベントを拾うための処理
        Dim i As Long
        For i = LBound(dayButtons) To UBound(dayButtons)
            dayButtons(i).Add Me("Label" & i)
        Next
    End With

End Sub
 
''
' 年が変更されたとき
'
Private Sub YearBox_Change()
    Call setclickedDates
End Sub

''
' 月が変更されたとき
'
Private Sub MonthBox_Change()
    Call setclickedDates
End Sub

''
' カレンダーの編集
'
Private Sub setclickedDates()
    
    With Me
        '年か月どちらか入ってなければ中止
        If .YearBox = "" Or .MonthBox = "" Then Exit Sub
    End With
    
    'ラベルの初期化
    Dim i As Long
    For i = 1 To LAST_BUTTON
        With Me("Label" & i)
            .Caption = ""
            .BackColor = Me.BackColor
        End With
    Next
    
    '選択年月を取得
    With Me
        Dim yy As Long: yy = .YearBox
        Dim mm As Long: mm = .MonthBox
    
        Dim firstDateOfMonth As Date
        firstDateOfMonth = DateSerial(yy, mm, 1)
    End With
    
    'カレンダーに日付を表示
    Dim n As Long
    Dim endDateOfMonth As Long
    n = Weekday(firstDateOfMonth) - 1
    endDateOfMonth = Day( _
        DateAdd("d", -1, DateAdd("m", 1, firstDateOfMonth)) _
    )
    
    Dim j As Long
    For j = 1 To endDateOfMonth
        With Me("Label" & j + n)
            .Caption = j '日
            
            'TextBoxの日付のみ色付け
            If DateSerial(yy, mm, j) <> clickedDate Then GoTo Continue
            .BackColor = RGB(255, 217, 102)
'            .ForeColor = RGB(255, 255, 255)
        End With
Continue:
    Next j

End Sub

''
' ひと月戻る
'
Private Sub SpinButton1_SpinUp()
  
  '1月のみ年月を調整
  With Me
    If .MonthBox.Value = 1 Then
      .YearBox.Value = .YearBox.Value - 1
      .MonthBox.Value = 12 '12月へ

    Else
      .MonthBox.Value = .MonthBox.Value - 1

    End If
  End With
  
End Sub
 
''
' ひと月進む
'
Private Sub SpinButton1_SpinDown()
  
    '12月のみ年月を調整
    With Me
        If .MonthBox.Value = 12 Then
            .YearBox.Value = .YearBox.Value + 1
            .MonthBox.Value = 1
        
        Else
            .MonthBox.Value = .MonthBox.Value + 1
        
        End If
    End With

End Sub
