VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "MessageRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "���R�[�h�̃R���N�V�������Ǘ����܂��B"

Option Explicit

Private pItems As Collection
Private pRecordId As Long

Property Get Items() As Collection
    Set Items = pItems
End Property

''
' �R���X�g���N�^
'
Private Sub Class_Initialize()
    Set pItems = New Collection
    pRecordId = 1

End Sub

''
' ���R�[�h��ǉ�����
'
' @param {MessageRecord} ���R�[�h1�s��
'
Public Function Add(ByVal record As MessageRecord) As MessageRecord

    pItems.Add record ', CStr(record.RecordId)

    Set Add = record

End Function

''
' ���R�[�h���擾����
'
' @param {Long} Item�̃L�[�i���l�j
'
' @return {MessageRecord} ���R�[�h1�s��
'
Public Function Item(ByVal key As Long) As MessageRecord
    Set Item = pItems.Item(key)
End Function

''
' ���R�[�h�Z�b�g��񎟌��z��ɒu��������
' �y�Q�Ɓz
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @return {array} [arr(0,0)]���̔z��
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
        '1�ł��l�������False��Ԃ�
        If IsEmpty(targets(i)) = False Then
            ret = False
            Exit Function
        End If
    Next

    isAllEmpty = ret

End Function