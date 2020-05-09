VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Array2d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = True
Attribute VB_Exposed         = False
Attribute VB_Description     = "�񎟌��z��������₷�����܂��B"
Option Explicit

' ���������������������� ���W���[���ϐ� ����������������������

Private pArray1d As Array1d
Private pArray1ds As Collection

Private pValues As Variant

' ���������������������� �R���X�g���N�^�� ����������������������

''
' ������
'
Private Sub Class_Initialize()

    Set pArray1ds = New Collection

End Sub

''
' �V�[�g�̒l����荞��
'
' @param {Worksheet} srcSheet ���[�N�V�[�g
'
' @return {Array2d} Me�i���\�b�h�`�F�[�����\�ɂ��邽�߁j
'
Public Function Import(ByRef srcSheet As Worksheet) As Array2d

    Me.Values = srcSheet.UsedRange.Value
    Call convertRowsToCollection(srcSheet)

    Set Import = Me

End Function

''
' �V�[�g�̒l��Array1d�N���X�̃R���N�V�����ɕϊ�����
'
' @param {Worksheet} srcSheet ���[�N�V�[�g
'
Private Sub convertRowsToCollection(ByRef srcSheet As Worksheet)

    Dim dataRange As Range
    Set dataRange = srcSheet.UsedRange

    Dim i As Long
    For i = 1 To dataRange.Rows.Count
        Dim arr1d As Array1d: Set arr1d = New Array1d
        arr1d.arr = dataRange.Rows(i).Value
        Call Me.Array1ds.Add(arr1d)
    Next i

End Sub

' ���������������������� �v���p�e�B ����������������������

''
' �V�[�g�̒l
'
Public Property Get Values() As Variant
Attribute Values.VB_Description = "�V�[�g�̒l��\���񎟌��z��"
Attribute Values.VB_UserMemId   = 0
    Values = pValues
End Property
Public Property Let Values(ByVal aValues As Variant)
Attribute Values.VB_Description = "�V�[�g�̒l��\���񎟌��z��"
Attribute Values.VB_UserMemId   = 0
    pValues = aValues
End Property

''
' Array1d�̃R���N�V����
'
Public Property Get Array1ds() As Collection
    Set Array1ds = pArray1ds
End Property
Public Property Set Array1ds(ByVal aArray1d As Collection)
    Set pArray1ds = aArray1ds
End Property

''
' �s��
'
Public Property Get RowCounts() As Long
    RowCounts = Array1ds.Count
End Property

''
' ��
'
Public Property Get ColCounts() As Long
    ColCounts = Me.Length(2)
End Property

''
' �z��Ƃ��Ă̗v�f��
'
Public Property Get Length(ByVal dimension As Long) As Long
    Length = UBound(Me.Values, dimension)
End Property

' ���������������������� ���\�b�h ����������������������

''
' 2�w�ɂȂ����z���񎟌��z��ɒu��������
'
' �y�Q�Ɓz
' https://qiita.com/11295/items/7364a80814bca5b734ff
'
' @param {array} �W���O�z��iarr(0)(0)�`���j
'
' @return {array} �񎟌��z��iarr(0, 0)�`���j
'
Public Function NestToFlat(ByRef nestArray2d As Variant) As Variant

    Dim flatArray2d() As Variant
    ReDim flatArray2d(0 To UBound(nestArray2d), _
                      0 To UBound(nestArray2d(0)))

    Dim r As Long
    Dim rowData As Variant
    For r = 0 To UBound(nestArray2d)
        Dim c As Long
        Dim cellValue As Variant
        For c = 0 To UBound(rowData)
            flatArray2d(r + 1, c + 1) = cellValue
        Next
    Next

    NestToFlat = flatArray2d

End Function