Attribute VB_Name = "GetDataType"
Option Explicit

Enum eDataType
    edSOMETHING = 0
    edHEADER_RawData
    edDate
    edHEADER_Image
    edHEADER_Date
    edHEADER_UserId
    edUserId
    edHEADER_UserName
    edUserName
    edHEADER_body
    edLine
End Enum

Sub GetDataType()

    Dim pi As New PerformanceImprovement
    pi.WaitingAnime

    'データテーブルの初期化
    Dim dataTypeList As ListObject
    Set dataTypeList = shDataType.ListObjects(1)

    If Not (dataTypeList.DataBodyRange Is Nothing) Then
        dataTypeList.DataBodyRange.Delete
    End If

    '最終行を取得
    Dim lastRow As Long
    With shRawData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With

    Dim valueType As eDataType
    valueType = 0

    Dim dataTypeSheetValues() As Variant
    ReDim dataTypeSheetValues(lastRow - 1)

    '降順に走査
    Dim i As Long
    Dim dataTypeListRow As Long: dataTypeListRow = 2
    For i = lastRow To 1 Step -1

        'ステータスバー
        pi.WaitingAnime (dataTypeListRow)

        Dim val As Variant
        val = shRawData.Cells(i, 1).Value

        Select Case val
            Case "▼ここに貼り付けてください。"
                valueType = edHEADER_RawData

            Case "画像が送信されました"
                valueType = edHEADER_Image

            Case "メッセージが届きました"
                valueType = edHEADER_Date

            Case "ID"
                valueType = edHEADER_UserId

            Case "ユーザ名"
                valueType = edHEADER_UserName

            Case "本文"
                valueType = edHEADER_body

            Case Else
                Select Case valueType
                    Case edHEADER_Date
                        If VarType(val) = vbString Then
                            If InStr(val, "kaguyaアプリ") = 0 _
                           And InStr(val, ":") = 0 Then

                                valueType = edLine
                            Else
                                valueType = edDate
                            End If

                        Else
                            valueType = edDate
                        End If

                    Case edDate
                        valueType = edLine

                    Case edHEADER_UserId
                        valueType = edHEADER_Date

                    Case edHEADER_UserName
                        valueType = edUserId

                    Case edHEADER_body
                        valueType = edUserName

                    Case Else
                        valueType = edLine
                End Select

        End Select

        'リストに追加
        Dim rowValues As Variant
        rowValues = Array(i, val, valueType)
        dataTypeSheetValues(dataTypeListRow - 2) = rowValues

        dataTypeListRow = dataTypeListRow + 1
    Next

    'シートに貼付できるように変換
    dataTypeSheetValues = get2dValues(dataTypeSheetValues)

    '空行があれば置き換え
    For dataTypeListRow = 0 To UBound(dataTypeSheetValues, 1)
        If dataTypeListRow = UBound(dataTypeSheetValues, 1) Then Exit For

        If dataTypeSheetValues(dataTypeListRow, 2) <> eDataType.edDate Then GoTo continue:
        If dataTypeSheetValues(dataTypeListRow + 1, 1) <> "" Then GoTo continue:

        dataTypeSheetValues(dataTypeListRow + 1, 2) = 0
continue:
    Next

    '表に出力
    shDataType.Cells(2, 1).Resize(UBound(dataTypeSheetValues, 1) + 1, _
                                  UBound(dataTypeSheetValues, 2) + 1).Value = dataTypeSheetValues

    '昇順に戻す
    Dim list As ListObject
    Set list = shDataType.ListObjects(1)
    list.Range.Sort Key1:=shDataType.Cells(1, 1), _
                    Order1:=xlAscending, _
                    Header:=xlYes, _
                    Orientation:=xlTopToBottom, _
                    SortMethod:=xlPinYin

End Sub

''
' 2層になった配列を二次元配列に置き換える
' 【参照】
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @param {array} [arr(0)(0)]式の配列
'
' @return {array} [arr(0,0)]式の配列
'
Private Function get2dValues(ByRef nest2dArr As Variant) As Variant

    Dim ret() As Variant
    ReDim ret(0 To UBound(nest2dArr), 0 To UBound(nest2dArr(0)))

    Dim r As Long: r = 0
    Dim rowData As Variant
    For Each rowData In nest2dArr
        Dim c As Long: c = 0
        Dim element As Variant
        For Each element In rowData
            ret(r, c) = element
            c = c + 1
        Next
        r = r + 1
    Next

    get2dValues = ret
End Function
