Attribute VB_Name = "FormatSheet"
Option Explicit

Sub completionWorksheets()

    With shDataType
        .Activate
        .Cells(2, 3).Select
    End With

    With shOutput
        .Activate
        .Cells(2, 5).Select
    End With

    Dim colorRange As Range
    With shOutput
        Set colorRange = .Range(.Columns(1), _
                                .Columns(.ListObjects(1).ListColumns.Count))
    End With

    Dim mColors As New MyColors
    Dim autoColor As New AutoColoring

    With autoColor
        .DeleteAllConditionsIn shOutput.Cells

        .setRowColors searchValues:=Array(True), _
                      ColoringRange:=colorRange, _
                      searchColNumber:=6, _
                      interiorColor:=mColors.LightOrange, _
                      fontColor:=mColors.Black, _
                      isItalic:=False, _
                      isBold:=False

        .drawRowBorder drawingRange:=colorRange, _
                       criteriaColNumber:=3, _
                       lineColor:=mColors.Black

        Const DOUBLE_QUATE As String = """"
        .IsSatisfied condition:="$E1=" & DOUBLE_QUATE & "EMPTY" & DOUBLE_QUATE, _
                     ColoringRange:=shOutput.Columns(5), _
                     fontColor:=mColors.LightBlue, _
                     isItalic:=True

        .IsSatisfied condition:="$F1=" & False, _
                     ColoringRange:=shOutput.Columns(6), _
                     fontColor:=mColors.LighterBlue, _
                     isItalic:=True
    End With

End Sub
