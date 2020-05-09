VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Array1d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "�z��������₷�����܂��B"
Option Explicit

' ���������������������� ���W���[���ϐ� ����������������������

Private pArray As Variant

' ���������������������� �R���X�g���N�^�� ����������������������

''
' �R���X�g���N�^
'
' @param {Long} maxIndex �C���f�b�N�X�̍ő�l
'
' @return {Array1d} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Init(Optional ByVal maxIndex As Long = 0) As Array1d

    ReDim pArray(maxIndex) '�z��̏�����

    Set Init = Me

End Function

' ���������������������� �v���p�e�B ����������������������

''
' �z��
'
Public Property Get arr() As Variant
    arr = pArray
End Property
Public Property Let arr(ByVal aArray As Variant)
    pArray = aArray
End Property
Public Property Set arr(ByVal aArray As Variant)
    Set pArray = aArray
End Property

''
' �v�f
'
Public Property Get Elements(ByVal Index As Long) As Variant
    Elements = pArray(Index)
End Property
Public Property Let Elements(ByVal Index As Long, _
                             ByVal assginValue As Variant)
    pArray(Index) = assginValue
End Property
Public Property Set Elements(ByVal Index As Long, _
                             ByRef assginValue As Variant)
    Set pArray(Index) = assginValue
End Property

''
' �v�f��
'
Public Property Get Length() As Long
    Length = UBound(pArray)
End Property

' ���������������������� ���\�b�h ����������������������

''
' �z����g������
'
' @param {Long} maxIndex �z��̗v�f��
'
' @return {Array1d} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Extend(Optional ByVal maxIndex As Long = 0) As Array1d

    '�g���T�C�Y���m�F
    Dim extension As Long
    If maxIndex = 0 Then '������@�c�H
        extension = UBound(pArray) + 1

    Else
        extension = maxIndex

    End If

    '�l��ێ������܂܊g��
    ReDim Preserve pArray(extension)

    Set Extend = Me

End Function

''
' �Ō���ɗv�f��ǉ�
'
' �y�Q�ƁzVBA ���I�z��ŗv�f�����ӎ������Ɋg��������@ | t-hom�fs diary
' https://thom.hateblo.jp/entry/2015/03/19/213019
'
' @param {Variant} ����l
'
' @return {Array1d} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Push(ByVal var As Variant) As Array1d

    '�z����g��
    Me.Extend

    '�Ō���ɒǉ�
    pArray(UBound(pArray)) = var

    Set Push = Me

End Function

''
' �z��𔽓]����
'
' �y�Q�l�z
' VBA �z��̕��т𔽓]������
' https://www.tipsfound.com/vba/02021
'
' @return {Array1d} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Reverse() As Array1d

    Dim low As Long:   low = LBound(pArray)
    Dim high As Long: high = UBound(pArray)
    Dim arrLength As Long: arrLength = Me.Length

    '���]�����z����쐬
    Dim reversedArray() As Variant
    ReDim Preserve reversedArray(low To high)
    Dim i As Long
    For i = 0 To arrLength
        reversedArray(low + i) = pArray(high - i)
    Next

    '���̔z��ɑ��
    For i = 0 To arrLength
        pArray(low + i) = reversedArray(low + i)
    Next

    Set Reverse = Me

End Function

''
' �z���A������
'
' @param {string} ��؂蕶��
'
' @return {string} �A������������
'
Public Function JoinVia(Optional ByVal delimiter As String = ", ") As String

    Dim str As String: str = Join(pArray, delimiter)
    If str = "" Then str = "EMPTY"

    JoinVia = str

End Function

''
' �l�����O�o�͂���
'
' @param {string} ��؂蕶��
'
Public Sub PrintAll(Optional ByVal delimiter As String = ", ")

    Dim str As String: str = JoinVia(delimiter)

    Debug.Print "[" & str & "]"

End Sub