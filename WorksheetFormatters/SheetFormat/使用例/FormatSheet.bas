Attribute VB_Name = "FormatSheet"
Option Explicit

''
' スタイルを作成してテーブルに設定する
'
Public Sub FormatSheet()

    Dim myFormat As AutoFormat
    Set myFormat = New AutoFormat

    Call myFormat.Apply(ActiveSheet)

End Sub

''
' Excelの標準フォントを修正する
'
Public Sub ModifyExcelFont()

    Dim myFormat As AutoFormat
    Set myFormat = New AutoFormat

    Call myFormat.SetExcelFont(ActiveWorkbook)

End Sub
