Attribute VB_Name = "Array1D"
Option Explicit

Sub arrTest()

    Dim Arr() As String
    ReDim Arr(0) 'Index 0Ç≈èâä˙âª

    Dim arra As New PowerArray
    arra.init (Arr)

    Dim i As Long
    For i = 0 To 15
        arra.Push (i)
    Next

    Dim str As String
    str = arra.JoinVia(",")

    str = arra.JoinVia(vbCrLf)

    Debug.Print str

    arra.PrintAll (",")

End Sub
