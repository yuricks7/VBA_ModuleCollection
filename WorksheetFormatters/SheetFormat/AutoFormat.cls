VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "AutoFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "ワークシートの見た目を調整します。"
Option Explicit

'━━━━━━━━━━━ モジュール変数 ━━━━━━━━━━━

Private pFont As String
Private pFontSize As Long
Private pHeaderFontSize As Long

Private pVerticalAlignment As Long
Private pRowHeight As Double

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

        'スタイルを変更（2回目はエラー？）
        On Error GoTo CreateStyle
        targetList.TableStyle = NEW_STYLE_NAME
        On Error GoTo 0

        'ブックの「既定のスタイル」に設定
        targetBook.DefaultTableStyle = NEW_STYLE_NAME

        'クエリの設定を変更
        If .QueryTables.Count > 0 Then
            Call SetQueryProperties(targetList)
        End If

        'シート名を変更
        With targetList

            Dim newName As String
            If isDefalutName(.Name) Then
                newName = Application.InputBox( _
                                Prompt:="シート名を入力してください。", _
                                Title:="シート名確認中…", _
                                Default:="data" _
                            )
                .Name = newName
            Else
                newName = .Name

            End If

            On Error Resume Next
            targetSheet.Name = newName
            On Error GoTo 0

        End With

Named:
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

        '設定値の取得に使用
        Dim myFormat As FieldFormat: Set myFormat = New FieldFormat

        '全体の設定
        With .Cells
            With .Font
                .Name = Me.Font
                .Size = Me.FontSize
            End With

            .RowHeight = Me.RowHeight
            .VerticalAlignment = Me.VerticalAlignment
            .NumberFormatLocal = myFormat.DisplayFormat(eNumbers)
        End With

        'ヘッダーの設定
        Dim sheetHeader As Range: Set sheetHeader = .Rows(1)
        sheetHeader.Font.Size = Me.HeaderFontSize
        Call SetColorFormula(sheetHeader)

        '列幅の設定
        With .Columns
            .ColumnWidth = myFormat.Width(eGridPaper)
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

    Dim sheetColumns As Range: Set sheetColumns = targetSheet.Columns

    '1列ごとに設定を適用する
    Dim i As Long
    For i = 1 To sheetColumns.Count
        With sheetColumns(i)
            If .Cells(1, 1).Value = "" Then Exit Sub

            '列名を見て列の設定を行う
            Dim fFormat As FieldFormat: Set fFormat = New FieldFormat
            Set fFormat.Range = sheetColumns(i)
            fFormat.ApplyByHeader

        End With
    Next i

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

''
' テーブル名がデフォルト値のままか確認する
'
' @param {String} tableName テーブル名
'
' @return {Boolean} テーブル名が「テーブル～」であるか
'
Private Function isDefalutName(ByVal tableName As String) As Boolean

    Dim ret As Boolean: ret = False
    If tableName Like "テーブル" & "*" Then ret = True

    isDefalutName = ret

End Function
