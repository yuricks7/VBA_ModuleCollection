VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pItems As Collection
Private pRecordId As Long

Property Get Items() As Collection
    Set Items = pItems
End Property

''
' コンストラクタ
'
Private Sub Class_Initialize()
    Set pItems = New Collection
    pRecordId = 1

End Sub

''
' レコードを追加する
'
' @param {MessageRecord} レコード1行分
'
Public Function Add(ByVal record As MessageRecord) As MessageRecord

    pItems.Add record ', CStr(record.RecordId)

    Set Add = record

End Function

''
' レコードを取得する
'
' @param {Long} Itemのキー（数値）
'
' @return {MessageRecord} レコード1行分
'
Public Function Item(ByVal key As Long) As MessageRecord
    Set Item = pItems.Item(key)
End Function

''
' レコードセットを二次元配列に置き換える
' 【参照】
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @return {array} [arr(0,0)]式の配列
'
Public Function GetValues() As Variant

    Dim ret() As Variant
    Dim iCounts As Long
    iCounts = Me.Items.Count
    ReDim ret(0 To iCounts - 1, 0 To Me.Item(1).Count - 1)

    Dim pi As New PerformanceImprovement

    Dim i As Long
    Dim r As Long
    For i = 1 To iCounts
        pi.ProgressBar iCounts, i

        Dim targets() As Variant
        With Me.Item(i)
            targets = Array(.RecordId, .PostTime, .UserId, .UserName, .Body, .HasImage)
            If isAllEmpty(targets) Then GoTo continue:

            r = i - 1
            ret(r, 0) = .RecordId
            ret(r, 1) = .PostTime
            ret(r, 2) = .UserId
            ret(r, 3) = .UserName
            ret(r, 4) = .Body
            ret(r, 5) = .HasImage
        End With

        If Me.Item(i) Is Nothing Then GoTo break:

continue:
    Next

break:
    GetValues = ret

End Function

Private Function isAllEmpty(ByRef targets As Variant) As Boolean

    Dim ret As Boolean
    ret = True

    Dim i As Long
    For i = 0 To Ubound(targets)
        '1つでも値があればFalseを返す
        If IsEmpty(targets(i)) = False Then
            ret = False
            Exit Function
        End If
    Next

    isAllEmpty = ret

End Function
