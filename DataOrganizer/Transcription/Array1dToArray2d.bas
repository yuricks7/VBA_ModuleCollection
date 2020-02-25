Attribute VB_Name = "FormatData"
Option Explicit

Sub FormatMessageData()

    Dim pi As New PerformanceImprovement
    pi.WaitingAnime

    Dim dataList As ListObject
    Dim dataListValues As Variant
    Set dataList = shDataType.ListObjects(1)
    dataListValues = dataList.DataBodyRange.Value

    '変数の初期化
    Dim recordCounts As Long: recordCounts = 1
    Dim record As New MessageRecord: record.Init (recordCounts) '出力1行分

    Dim records As MessageRecords: Set records = New MessageRecords

    'データの行数分走査
    Dim r As Long
    For r = 1 To UBound(dataListValues)
        'ステータスバー
        pi.WaitingAnime r

        '現在のデータ
        Dim dataValue As Variant: dataValue = dataListValues(r, 2)
        Dim dataType As eDataType: dataType = dataListValues(r, 3)

        'データ・タイプによって振り分け
        Select Case dataType
            Case eDataType.edHEADER_RawData

            Case eDataType.edDate

                If VarType(dataValue) = vbString Then
                    If InStr(dataValue, "kaguya") = 0 Then
                         dataType = edLine
                         GoTo pushLine:

                    Else
                        dataValue = TimeValue(Right(dataValue, 5)) '表記の統一

                    End If
                End If

                record.PostTime = dataValue

            Case eDataType.edHEADER_Image
                record.HasImage = True

            Case eDataType.edHEADER_Date
                '1件目は飛ばす
                If recordCounts = 1 And record.UserId = "" Then GoTo continue:
                records.Add record

                '再度初期化
                Set record = New MessageRecord
                recordCounts = recordCounts + 1
                record.Init (recordCounts)

            Case eDataType.edHEADER_UserId

            Case eDataType.edUserId
                record.UserId = dataValue

            Case eDataType.edHEADER_UserName

            Case eDataType.edUserName
                record.UserName = dataValue

            Case eDataType.edHEADER_body

            Case eDataType.edLine
pushLine:
                record.Lines.Push (dataValue)

        End Select

continue:
    Next

    '最後のレコード
    records.Add record

    Dim outputValues() As Variant
    ReDim outputValues(records.Items.Count - 1, 5)
    outputValues = records.GetValues

    Dim messageList As ListObject
    Set messageList = shOutput.ListObjects(1)

    Call setValues(outputValues, messageList)

End Sub

''
' シートに二次元配列を代入
'
' @param {variant} 二次元配列
' @param {listobject} 貼り付け先のテーブル
'
Sub setValues(ByRef source2dArray() As Variant, _
              ByRef destinationListObj As ListObject)

    Dim bodyRange As Range
    Set bodyRange = destinationListObj.DataBodyRange

    Dim upperLeft As Range
    If bodyRange Is Nothing Then
        Set upperLeft = destinationListObj.Parent.Cells(2, 1)
    Else
        Set upperLeft = bodyRange.Cells(1, 1)
    End If

    upperLeft.Resize(UBound(source2dArray, 1) + 1, _
                     UBound(source2dArray, 2) + 1).Value = source2dArray

End Sub
