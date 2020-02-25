Option Explicit

''
' テーブルをコピペ
'
Sub test1()

    Dim dataList As ListObject
    Set dataList = shDataType.ListObjects(1)

    Dim dataListValues As Variant
    dataListValues = dataList.DataBodyRange.Value

    Dim messageList As ListObject
    Set messageList = shOutput.ListObjects(1)

    ' Dim destRange As Range
    ' Set destRange = messageList.DataBodyRange

    With messageList
        .Range(.cells(1,1), _
               .cells(ubound(dataListValues, 1), ubound(dataListValues, 2)))
    End With

End Sub


''
' テーブルをコピペ
'
Sub test2()

    Dim dataList As ListObject
    Set dataList = shDataType.ListObjects(1)

    Dim dataListValues As Variant
    dataListValues = dataList.DataBodyRange.Value

    Dim messageList As ListObject
    Set messageList = shOutput.ListObjects(1)

    Call setValues(dataListValues, messageList)

End Sub

''
' 二次元配列を代入
'
' @param {variant} 二次元配列
' @param {listobject} 貼り付け先のテーブル
'
Sub setValues(ByRef source2dArray As Variant, _
              ByRef destinationListObj As ListObject)

    Dim bodyRange As Range
    Set bodyRange = destinationListObj.DataBodyRange

    Dim upperLeft As Range
    If bodyRange Is Nothing Then
        Set upperLeft = destinationListObj.Parent.Cells(2, 1)
    Else
        Set upperLeft = bodyRange.Cells(1, 1)
    End If

    Dim lowerRight As Range
    Set lowerRight = upperLeft.Cells(UBound(source2dArray, 1), UBound(source2dArray, 2))

    Dim targetRange As Range
    Set targetRange = Range(upperLeft, lowerRight)

    With targetRange
        .Value = source2dArray
    End With

    MsgBox "貼れましたよ。", vbInformation

End Sub
