VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArray() As Variant
Private pLength As Long
Private pDimension As Long

Property Get arr() As Variant
    arr = pArray
End Property

Property Set arr(ByRef aArray As Variant)
    pArray = aArray
End Property

Property Get Length() As Long
    Length = UBound(pArray)
End Property

''
' �R���X�g���N�^
'
' @param {variant} �����Ώۂ̔z��
'
' @return {object} Me
'
Function Init(Optional ByVal uBoundIndex As Long = 0) As PowerArray

    ReDim pArray(uBoundIndex) '�z��̏�����

    Set Init = Me
End Function

''
' �z��̎����������O�o�͂���
'
' �yExcel VBA�z�z��̎��������擾����֐��R�[�h | oki2a24  Just another oki2a24 �u���S
' https://oki2a24.com/2016/03/09/get-dimension-number-in-excel-vba/
'
' @param {variant} �z��
'
' @return {long} ������
'
Property Get Dimension() As Long
    Dim temp As Long
    temp = 1

    Dim forDecision As Long

    '���������m�F
    On Error Resume Next
    Do While Err.Number = 0 '�G���[���N������X�g�b�v
        forDecision = UBound(arr, temp)
        temp = temp + 1
    Loop

    '�l�̒���
    '�i�G���[�������Ɏ��� + 1 �ƂȂ��Ă���̂ŁA2�������j
    pDimension = temp - 2

    Dimension = pDimension
End Property

''
' �Ō���ɗv�f��ǉ�
'
' �y�Q�ƁzVBA ���I�z��ŗv�f�����ӎ������Ɋg��������@ | t-hom�fs diary
' https://thom.hateblo.jp/entry/2015/03/19/213019
'
' @param {variant} �ǉ�����l
'
' @return {variant} �ǉ���̔z��
'
Function Push(ByVal var As Variant) As Variant

    '�Ō���ɒǉ�
    pArray(UBound(pArray)) = var

    '�z����g��
    ReDim Preserve pArray(UBound(pArray) + 1)

    Push = pArray

End Function

''
' �w��̃C���f�b�N�X�ʒu�ɒl����
'
' @param {long} �Ώۂ̃C���f�b�N�X
' @param {variant} �ǉ�����l
'
' @return {variant} �ǉ���̔z��
'
Function Assign(ByVal targetIndex As Long, _
                ByVal var As Variant) As Variant

    pArray(targetIndex) = var

    Assign = pArray

End Function

''
' �z��𔽓]����
'
' @param {string} ��؂蕶��
'
' @return {string} �A������������
'
Function Reverse() As Variant

    Dim low As Long: low = LBound(pArray)
    Dim high As Long: high = UBound(pArray)

    Dim tempArray() As Variant
    ReDim Preserve tempArray(low To high)

    Dim arrLength As Long
    arrLength = high - low

    Dim i As Long
    For i = 0 To arrLength
        tempArray(low + i) = pArray(high - i)
    Next

    For i = 1 To arrLength
        pArray(low + i) = tempArray(low + i)
    Next

    Reverse = pArray
End Function

''
' 2�w�ɂȂ����z���񎟌��z��ɒu��������
' �y�Q�Ɓz
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @param {array} [arr(0)(0)]���̔z��
'
' @return {array} [arr(0,0)]���̔z��
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

''
' �z���A������
'
' @param {string} ��؂蕶��
'
' @return {string} �A������������
'
Function JoinVia(ByVal delimiter As String) As String

    Dim str As String
    str = Join(pArray, delimiter)

    If str = "" Then
        JoinVia = "EMPTY"
        Exit Function
    End If

    Dim Length As Long
    Length = Len(str)

    '�Ō�̋�؂蕶�����폜
    str = Left(str, Length - Len(delimiter))
    JoinVia = str

End Function

''
' ���s��t���Ȃ���l�����ׂĘA������
'
' @return {string} �A������������
'
Function JoinWithCRLF() As String
    '������
    Dim str As String: str = ""

    '2�ځ`�Ō��1�O�܂ŁA���s�ƌ���
    Dim a As Long
    For a = 1 To UBound(arr) - 1
        '1��
        If a = 0 Then
            str = pArray(0)

        '�Ō�͒l�̒ǉ��̂�
        ElseIf a = UBound(pArray) Then
            str = str & pArray(UBound(pArray))

        '���̑��́A�l+���s������
        Else
            str = str & vbCrLf & CStr(pArray(a))
        End If
    Next

    JoinWithCRLF = str
End Function

''
' �l�����O�o�͂���
'
' @param {string} ��؂蕶��
'
Sub PrintAll(ByVal delimiter As String)

    Dim p As String
    p = JoinVia(delimiter)

    Debug.Print "[" & p & "]"

End Sub