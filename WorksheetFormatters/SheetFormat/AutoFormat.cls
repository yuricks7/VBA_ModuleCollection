VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pFont As String
Private pFontSize As Long
Private pHeaderFontSize As Long

Private pVerticalAlignment As Long
Private pRowHeight As Double

Private pHorizontalAlignment As Long
Private pColumnWidth_MAX As Double
Private pColumnWidth_MIDDLE As Double
Private pColumnWidth_NARROW As Double

Private pColumnWidth_GridPaper As Double
Private pColumnWidth_MinutesAndHour As Double
Private pColumnWidth_Date_WithDayOfTheWeek As Double
Private pColumnWidth_DateTime_WithDayOfTheWeek As Double
Private pColumnWidth_Boolean As Double

Private pTypeOfStandard As String
Private pTypeOfString As String
Private pTypeOfNumbers As String
Private pTypeOfNumbers00 As String

Private pTypeOfDate As String
Private pTypeOfDateTime As String
Private pTypeOfDate_WithDayOfTheWeek As String
Private pTypeOfDateTime_WithDayOfTheWeek As String

'━━━━━━━━━━━ プロパティ ━━━━━━━━━━━

''
' フォント
'
Property Get Font() As String
    
    pFont = "Meiryo UI"
    Font = pFont

End Property

Property Get FontSize() As Long
    
    pFontSize = 8
    FontSize = pFontSize

End Property

Property Get HeaderFontSize() As Long
    
    pHeaderFontSize = 9
    HeaderFontSize = pHeaderFontSize

End Property

''
' 標準の垂直位置
'
Property Get VerticalAlignment() As Long
    
    pVerticalAlignment = xlBottom
    VerticalAlignment = pVerticalAlignment

End Property

Property Let VerticalAlignment(ByVal aVerticalAlignment As Long)

    '既定値の使い方あってる…？
    If IsEmpty(aVerticalAlignment) Then aVerticalAlignment = xlBottom
    pVerticalAlignment = aVerticalAlignment

End Property

''
' 行の高さ
'
Property Get RowHeight() As Double
    
    pRowHeight = 14.4 '24px
    RowHeight = pRowHeight

End Property

''
' 標準の水平位置
'
Property Get HorizontalAlignment() As Long
    
    pHorizontalAlignment = xlRight
    HorizontalAlignment = pHorizontalAlignment

End Property

Property Let HorizontalAlignment(ByVal aHorizontalAlignment As Long)

    '既定値の使い方あってる…？
    If IsEmpty(aHorizontalAlignment) Then aHorizontalAlignment = xlBottom
    pHorizontalAlignment = aHorizontalAlignment

End Property

''
' 列幅
'
Property Get ColumnWidth_MAX() As Double
    
    pColumnWidth_MAX = 46.88 '380px
    ColumnWidth_MAX = pColumnWidth_MAX

End Property

Property Get ColumnWidth_MIDDLE() As Double
    
    pColumnWidth_MIDDLE = 34.38 '280px
    ColumnWidth_MIDDLE = pColumnWidth_MIDDLE

End Property

Property Get ColumnWidth_NARROW() As Double
    
    pColumnWidth_NARROW = 14.38 '120px
    ColumnWidth_NARROW = pColumnWidth_NARROW

End Property

'Excel方眼紙
Property Get ColumnWidth_GridPaper() As Double
    
    pColumnWidth_GridPaper = 2.88 '28px
    ColumnWidth_GridPaper = pColumnWidth_GridPaper

End Property

'H列、M列
Property Get ColumnWidth_MinutesAndHour() As Double
    
    pColumnWidth_MinutesAndHour = 3.88 '36px
    ColumnWidth_MinutesAndHour = pColumnWidth_MinutesAndHour

End Property

Property Get ColumnWidth_Date_WithDayOfTheWeek() As Double
    
    pColumnWidth_Date_WithDayOfTheWeek = 14.38 '120px
    ColumnWidth_Date_WithDayOfTheWeek = pColumnWidth_Date_WithDayOfTheWeek

End Property

Property Get ColumnWidth_DateTime_WithDayOfTheWeek() As Double
    
    pColumnWidth_DateTime_WithDayOfTheWeek = 20.38 '168px
    ColumnWidth_DateTime_WithDayOfTheWeek = pColumnWidth_DateTime_WithDayOfTheWeek

End Property

'True / False
Property Get ColumnWidth_Boolean() As Double
    
    pColumnWidth_Boolean = 6.88 '60px
    ColumnWidth_Boolean = pColumnWidth_Boolean

End Property

''
' セルの表示形式
'
Property Get TypeOfStandard() As String
    
    pTypeOfStandard = "G/標準"
    TypeOfStandard = pTypeOfStandard

End Property

Property Get TypeOfString() As String
    
    pTypeOfString = "@"
    TypeOfString = pTypeOfString

End Property

Property Get TypeOfNumbers() As String
    
    pTypeOfNumbers = "#,##0_ ;[赤]▼ #,##0 "
    TypeOfNumbers = pTypeOfNumbers

End Property

Property Get TypeOfNumbers00() As String
    
    pTypeOfNumbers00 = "00"
    TypeOfNumbers00 = pTypeOfNumbers00

End Property

Property Get TypeOfDate() As String
    
    pTypeOfDate = "yyyy/mm/dd"
    TypeOfDate = pTypeOfDate

End Property

Property Get TypeOfDateTime() As String
    
    pTypeOfDateTime = "yyyy/mm/dd hh:mm"
    TypeOfDateTime = pTypeOfDateTime

End Property

Property Get TypeOfDate_WithDayOfTheWeek() As String

    pTypeOfDate_WithDayOfTheWeek = "yyyy/mm/dd (aaa)"
    TypeOfDate_WithDayOfTheWeek = pTypeOfDate_WithDayOfTheWeek

End Property

Property Get TypeOfDateTime_WithDayOfTheWeek() As String

    pTypeOfDateTime_WithDayOfTheWeek = "yyyy/mm/dd (aaa) hh:mm"
    TypeOfDateTime_WithDayOfTheWeek = pTypeOfDateTime_WithDayOfTheWeek

End Property

'━━━━━━━━━━━ メソッド ━━━━━━━━━━━

''
' Excelの標準フォントを変える
'
' 【参考】
' 【Excel VBA】ブックのフォント、まとめて変更しませんか？ | ゆるおたノート
' https://www.yuru-wota.com/entry/VBA/set-favorite-font
'
Public Sub SetExcelFont(ByRef targetBook As Workbook)

    With Me
        Dim strFontName As String: strFontName = .Font
        Dim numFontSize As Long:   numFontSize = .FontSize
    End With

    'Excelアプリケーション自体のフォントを変更
    With Application
        .StandardFont = strFontName
        .StandardFontSize = numFontSize
    End With

    Dim ws As Worksheet
    For Each ws In targetBook.Worksheets
        With ws
            'シートの標準フォントを統一
            With .Cells.Font
                    .Name = strFontName
                    .Size = numFontSize
            End With
            
            '文字の配置を調整
            .Rows.RowHeight = Me.RowHeight
            .Cells.VerticalAlignment = Me.VerticalAlignment
        End With
    Next ws

    Dim msg As String
    Dim strFontSize As String: strFontSize = CStr(numFontSize)

    msg = "設定を変更しました。"
    msg = msg & vbCrLf & "----------------------------"
    msg = msg & vbCrLf & "▼Excelアプリケーション"
    msg = msg & vbCrLf & "フォント：" & strFontName
    msg = msg & vbCrLf & "サイズ：" & strFontSize & " px"
    msg = msg & vbCrLf & "（ダウンロードしたファイルには反映されません。）"
    msg = msg & vbCrLf & "----------------------------"
    msg = msg & vbCrLf & "▼ワークシート"
    msg = msg & vbCrLf & "フォント：" & strFontName
    msg = msg & vbCrLf & "----------------------------"

    MsgBox msg

    MsgBox "アプリケーションの設定を変更した場合、" & vbCrLf & _
                "すぐに反映が必要な時は一旦Excelを再起動してください。"

End Sub

''
' テーブルとシートの設定を行う
'
Public Sub Apply(ByRef targetSheet As Worksheet)
    
    Dim targetBook As Workbook: Set targetBook = targetSheet.Parent
    If targetSheet Is Nothing Then Set targetSheet = targetBook.Worksheets(1)

    Const NEW_STYLE_NAME As String = "lightBlueBackStyle"

    With targetSheet
        Dim targetList As ListObject
        On Error GoTo NoTable
            Set targetList = .ListObjects(1)
        On Error GoTo 0
        
        'スタイルを変更（2回目はエラーになるような？）
        On Error GoTo CreateStyle
        targetList.TableStyle = NEW_STYLE_NAME
        On Error GoTo 0
        
        'ブックの「既定のスタイル」に設定
        targetBook.DefaultTableStyle = NEW_STYLE_NAME
        
        'クエリの設定を変更
        If .QueryTables.Count > 0 Then
            Call SetQueryProperties(targetList)
        End If
        
        'シート名をテーブル名と同期
        On Error Resume Next 'エラーならデフォルトのまま
        targetSheet.Name = targetList.Name
        On Error GoTo 0
    End With
    
    
    '▼後始末
    Call SetSheetStyle(targetSheet)
    
    MsgBox "テーブルを整形しました。", vbInformation
    
    Exit Sub

NoTable:
    MsgBox "選択中のシートにテーブルがありません。"
    MsgBox "処理を中止します｡ "
    Exit Sub

CreateStyle:
        Call createTableStyle(NEW_STYLE_NAME, targetBook)
        Resume 'エラーが発生したステートメントに戻る

End Sub

''
' テーブルのスタイルを追加する
'
' @param {String}   newStyleName
' @param {Workbook} targetBook
'
Private Sub createTableStyle(ByVal newStyleName As String, _
                             ByRef targetBook As Workbook)
        
    'スタイルの追加
    Dim newStyle As TableStyle
    Set newStyle = targetBook.TableStyles.Add(newStyleName)
    
    Call changeTableDesign(newStyle)
        
End Sub

''
' テーブルスタイルの設定を変更する
'
' 【参考】
' 【Excel VBA】テーブル変換とスタイル変更 - テーブル変換するところまで標準モジュールで書いてみる | ゆるおたノート
' https://www.yuru-wota.com/entry/VBA/set-table-safely
'
' @param {TableStyle} tableStyleObj 設定に使用するTableStyleオブジェクト
'
Private Sub changeTableDesign(ByRef tableStyleObj As TableStyle)
    
    'テーブル全体（WholeStyle）
    Dim black As Long:         black = RGB(0, 0, 0)
    Dim lightGray As Long: lightGray = RGB(208, 206, 206)
    
    Call setWholeStyle(tableStyleObj, black, lightGray)
    
    '見出し行（HeaderRowStyle）
    Dim deepBlue As Long: deepBlue = RGB(0, 32, 96)
    Dim white As Long:       white = RGB(255, 255, 255)

    Call setHeaderStyle(tableStyleObj, deepBlue, white, True)

End Sub

''
' テーブル全体の設定
'
' @param {TableStyle} tableStyleObj  設定に使用するTableStyleオブジェクト
' @param {Long}       outerLineColor 外枠の色
' @param {Long}       innerLineColor 外枠の色
'
Private Sub setWholeStyle(ByRef tableStyleObj As Variant, _
                          Optional ByVal outerLineColor As Long = 0, _
                          Optional ByVal innerLineColor As Long = 13553360)

    Dim wholeTableElements As Variant
    Set wholeTableElements = tableStyleObj.TableStyleElements(xlWholeTable)
    
    Dim outerLineConstants As Variant
    outerLineConstants = Array( _
                             xlEdgeTop, _
                             xlEdgeBottom, _
                             xlEdgeLeft, _
                             xlEdgeRight _
                         )
    
    Call setLines( _
        wholeTableElements, _
        outerLineConstants, _
        outerLineColor, _
        xlContinuous, _
        xlMedium _
    )
    
    Dim innerLineConstants As Variant
    innerLineConstants = Array( _
                             xlInsideVertical, _
                             xlInsideHorizontal _
                         )
    
    Call setLines( _
        wholeTableElements, _
        innerLineConstants, _
        innerLineColor, _
        xlContinuous, _
        xlThin _
    )

End Sub

''
' 罫線を設定する
'
' @param {Variant} StyleElements 設定するStyleElementsオブジェクト
' @param {Variant} linePositions 線の場所を表す組み込み定数の配列
' @param {Long} targetColor      線の色
' @param {Long} lineStyle        [Optional] 線の種類
' @param {Long} thickness        [Optional] 線の太さ
'
Private Sub setLines(ByRef StyleElements As Variant, _
                     ByRef linePositions As Variant, _
                     ByVal targetColor As Long, _
                     Optional ByVal lineStyle As Long = xlContinuous, _
                     Optional ByVal thickness As Long = xlThin)
 
    With StyleElements
        Dim i As Long
        For i = 0 To UBound(linePositions)
            With .Borders(linePositions(i))
                .Color = targetColor
                .lineStyle = lineStyle
                .Weight = thickness
            End With
        Next i
    End With

End Sub

''
' ヘッダーの設定
'
' @param {TableStyle} tableStyleObj 設定するTableStyleオブジェクト
' @param {Long} interiorColor       [Optional] 内容
' @param {Long} fontColor           [Optional] 内容
' @param {Long} isBold              [Optional] 内容
'
Private Sub setHeaderStyle(ByRef tableStyleObj As TableStyle, _
                           Optional ByVal interiorColor As Long = 16777215, _
                           Optional ByVal fontColor As Long = 6299648, _
                           Optional ByVal isBold As Boolean = True)

   Dim headerRowElements As Variant
   Set headerRowElements = tableStyleObj.TableStyleElements(xlHeaderRow)
   
    With headerRowElements
        .Interior.Color = interiorColor
        
        With .Font
            .Color = fontColor
            .Bold = isBold
        End With
    
    End With

End Sub

''
' クエリテーブルの設定
'
' @param {ListObject} targetList 対象のテーブル
'
Public Sub SetQueryProperties(ByRef targetList As ListObject)
    
    With targetList.QueryTable
        .RowNumbers = False           '行番号
        .FillAdjacentFormulas = False 'クエリ テーブルの右側の数式を自動的に更新
        .PreserveFormatting = True    '書式を保持する
        .RefreshOnFileOpen = False    'ファイルを開いたら更新する
        .BackgroundQuery = True       'バックグラウンドで更新の許可
        .RefreshStyle = xlInsertDeleteCells '新しいレコードセットと行数を正確に一致させる
        .SavePassword = False         'ODBC接続文字列のパスワード情報を保存
        .SaveData = True              'ピボットテーブル・レポートの定義とデータを保存
        .AdjustColumnWidth = False    '列幅を自動調整
        .RefreshPeriod = 0            'クエリ・テーブルを自動更新する間隔
        .PreserveColumnInfo = True    'クエリ・テーブルが更新されるごとに、
        '列の並べ替え､フィルター設定､レイアウトに関する情報を保存
    End With

End Sub

''
' データシートの初期設定
'
' @param {Worksheet} targetSheet 対象のシート
'
Public Sub SetSheetStyle(ByRef targetSheet As Worksheet)

    With targetSheet
        Dim navyBlue As Long: navyBlue = RGB(0, 32, 96)
        .Tab.Color = navyBlue
        
        '全体の設定
        With .Cells
            With .Font
                .Name = Me.Font
                .Size = Me.FontSize
            End With
            
            .RowHeight = Me.RowHeight
            .VerticalAlignment = Me.VerticalAlignment
            .NumberFormatLocal = Me.TypeOfNumbers
        End With
        
        'ヘッダーの設定
        Dim sheetHeader As Range: Set sheetHeader = .Rows(1)
        sheetHeader.Font.Size = Me.HeaderFontSize
        Call SetColorFormula(sheetHeader)
        
        '列幅の設定
        With .Columns
            .ColumnWidth = Me.ColumnWidth_GridPaper
            .AutoFit
        End With
        
        '列ごとの設定
        Call SetColumnStyles(targetSheet)
        
        '後処理
        sheetHeader.HorizontalAlignment = xlLeft
        
        .Cells(2, 2).Select
        ActiveWindow.FreezePanes = True '既に設定されていたら変わらない
        
    End With

End Sub

''
' 列の設定
'
' @param {Worksheet} 対象のシート
'
Public Sub SetColumnStyles(ByRef targetSheet As Worksheet)

    '列の設定
    With targetSheet
        Dim i As Long
        For i = 1 To .Columns.Count
            With .Columns(i)
                If .Cells(1, 1).Value = "" Then Exit Sub
            
                '列名を見て列の設定を行う
                Dim isApplied As Boolean
                isApplied = applyColumnSettingsByHeader(.Columns(i))
                
                'その他
                If isApplied = False Then Call setColumnWidth(.Columns(i))
                
            End With
        Next i
    End With

End Sub

''
' 列の設定値を適用
'
' @param {Range} targetColumn 適用先
'
Private Function applyColumnSettingsByHeader(ByRef targetColumn As Range) As Boolean

    Dim headerString As String
    headerString = targetColumn.Cells(1, 1).Value

    Dim left2 As String:   left2 = Left(headerString, 2)
    Dim right1 As String: right1 = Right(headerString, 1)
    Dim right2 As String: right2 = Right(headerString, 2)
    Dim right3 As String: right3 = Right(headerString, 3)
    Dim right4 As String: right4 = Right(headerString, 4)

    With Me
        '時刻（H列、M列）
        Select Case left2
            Case "H_", "時_"
                Call applyColumnSettings(targetColumn, xlRight, .TypeOfNumbers, .ColumnWidth_MinutesAndHour)
                applyColumnSettingsByHeader = True
                Exit Function
            
            Case "M_", "分_"
                Call applyColumnSettings(targetColumn, xlLeft, .TypeOfNumbers00, .ColumnWidth_MinutesAndHour)
                applyColumnSettingsByHeader = True
                Exit Function
        
        End Select
        
        Select Case True '早期リターン
            '日時
            Case right2 = "日時"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfDateTime_WithDayOfTheWeek, .ColumnWidth_DateTime_WithDayOfTheWeek)
                applyColumnSettingsByHeader = True
                Exit Function

            Case right1 = "日"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfDate_WithDayOfTheWeek, .ColumnWidth_Date_WithDayOfTheWeek)
                applyColumnSettingsByHeader = True
                Exit Function
            
            'ID
            Case right2 = "ID"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
        
            'No
            Case (right2 = "No"), (right3 = "No.")
                Call applyColumnSettings(targetColumn, xlRight, .TypeOfNumbers, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
        
            'カテゴリ
            Case (right2 = "カテ"), (right4 = "カテゴリ")
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
            
            'True/False
            Case (left2 = "is"), (right1 = "？"), (right1 = "?")
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, .ColumnWidth_Boolean)
                applyColumnSettingsByHeader = True
                Exit Function
                                        
        End Select

    End With

End Function

''
' 列の設定を適用する
'
' @param {Range}   targetColumn         該当列
' @param {Long}    aHorizontalAlignment 値の水平位置
' @param {String}  aNumberFormatLocal   値の書式
' @param {Long}    aColumnWidth         列幅
' @param {Boolean} canAutofit           幅を自動調整してよいか
'
Private Sub applyColumnSettings(ByRef targetColumn As Range, _
                                Optional ByVal aHorizontalAlignment As Long = xlRight, _
                                Optional ByVal aNumberFormatLocal As String, _
                                Optional ByVal aColumnWidth As Long, _
                                Optional ByVal canAutofit As Boolean = False)
    
    With targetColumn
        .HorizontalAlignment = aHorizontalAlignment
        
        If aNumberFormatLocal = "" Then aNumberFormatLocal = Me.TypeOfStandard
        .NumberFormatLocal = aNumberFormatLocal
        
        If canAutofit = False Then
            If IsEmpty(aColumnWidth) Then aColumnWidth = Me.ColumnWidth_NARROW
            .ColumnWidth = aColumnWidth
        
        Else
            .AutoFit
        
        End If
    
    End With

End Sub

''
' 列を既定の幅に調える
'
' @param {Range} targetColumn 幅を変える列
'
Private Sub setColumnWidth(ByRef targetColumn As Range)

    With targetColumn

        Select Case .ColumnWidth
            Case Is > Me.ColumnWidth_MAX
                .ColumnWidth = Me.ColumnWidth_MAX
            
            Case Is > Me.ColumnWidth_MIDDLE
                .ColumnWidth = Me.ColumnWidth_MIDDLE
            
            Case Is > Me.ColumnWidth_MIDDLE
                .ColumnWidth = Me.ColumnWidth_NARROW
            
        End Select
        
    End With

End Sub

''
' 条件付き書式の設定
' （関数がある列はヘッダーに色付けする）
'
' @param {Range} targetRow 設定の対象範囲
'
Public Sub SetColorFormula(ByRef targetRow As Range)
    
    With targetRow
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:="=ISFORMULA(A$2)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With targetRow.FormatConditions(1)
            With .Font
                .Bold = True
                .Color = RGB(0, 0, 0) '黒
            End With
            
            .Interior.Color = RGB(226, 239, 218) '淡い緑
            
            .StopIfTrue = False
        End With
    End With

End Sub
