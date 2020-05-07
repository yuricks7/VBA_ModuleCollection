VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "FsoZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Option Explicit

Private pFso As Object

Private Sub Class_Initialize()

    Set pFso = CreateObject("Scripting.FileSystemObject")

End Sub

''
' �t�@�C���E�t�H���_��ZIP�`���ň��k
'
' �y�Q�l�z
' ZIP�`���ň��k�E�𓀂��s��VBA�}�N�� | ���S�Ҕ��Y�^
' https://www.ka-net.org/blog/?p=7605
'
' @param {Variant} srcPath        ���t�@�C���E�t�H���_
' @param {Variant} dstFolderPath  �o�͐�i�w�肵�Ȃ��ꍇ�͌��t�@�C���E�t�H���_�Ɠ����ꏊ�j
'
' @return {Boolean} ���k�̐���
'
Public Function Zip(ByVal srcPath As Variant, _
                    Optional ByVal dstFolderPath As Variant = "") As Boolean

    On Error GoTo ErrorCatch

    With pFso
        '�p�X�̎�ނ��`�F�b�N
        If isFolder(dstFolderPath) Then GoTo CreateZipFile

        Select Case True
            Case isFolder(srcPath)
                dstFolderPath = srcPath

            Case isFile(srcPath)
                dstFolderPath = .GetFile(srcPath).ParentFolder.Path

            Case Else
                Exit Function

        End Select

CreateZipFile:
        '�o�͐�̐ݒ�
        Dim dstFilePath As Variant
        dstFilePath = addPathSeparator(dstFolderPath) & .GetBaseName(srcPath) & ".zip"

        '���ZIP�t�@�C�����쐬
        With .CreateTextFile(dstFilePath, True)
          .Write ChrW(&H50) & ChrW(&H4B) & ChrW(&H5) & ChrW(&H6) & String(18, ChrW(0))
          .Close
        End With

    End With

    'ZIP�t�@�C���Ɍ��t�@�C���̃R�s�[���i�[
    '�i�t�H���_���t�H���_���R�s�[�����j
    With CreateObject("Shell.Application")
        With .Namespace(dstFilePath)
'�����Ȃ�
'            Application.DisplayAlerts = False
            .CopyHere srcPath
'            Application.DisplayAlerts = True

            Do While .Items.Count < 1
                DoEvents '�V�X�e���ɐ����n��
            Loop
        End With
    End With

    Zip = True
    Exit Function

ErrorCatch:

    MsgBox "���k�𐋍s�ł��܂���ł����B"

    Zip = False

End Function

''
' ZIP�t�@�C������
'
' �y�Q�l�z
' ZIP�`���ň��k�E�𓀂��s��VBA�}�N�� | ���S�Ҕ��Y�^
' https://www.ka-net.org/blog/?p=7605
'
' @param {Variant} srcZipPath    Zip�t�@�C���̃p�X
' @param {Variant} dstFolderPath �W�J��i�w�肵�Ȃ��ꍇ�͌��t�@�C���E�t�H���_�Ɠ����ꏊ�j
'                                ���W�J��ɓ����t�@�C�����������ꍇ�̓��[�U�[���f�ŏ���
'
' @return {Boolean} �𓀂̐���
'
Public Function UnZip(ByVal srcZipPath As Variant, _
                      Optional ByVal dstFolderPath As Variant = "") As Boolean

    '���s�O�`�F�b�N
    With pFso
        If .FileExists(srcZipPath) = False Then GoTo ErrorCatch '�t�@�C���̑��݊m�F
        If LCase(.GetExtensionName(srcZipPath)) <> "zip" Then GoTo ErrorCatch '�g���q

        '�t�@�C�����w�肵�Ă���ꍇ�́A���̐e�t�H���_�ɕύX
        If isFolder(dstFolderPath) = False Then
            dstFolderPath = .GetFile(srcZipPath).ParentFolder.Path
        End If
    End With

    '��
    With CreateObject("Shell.Application")
        .Namespace(dstFolderPath).CopyHere .Namespace(srcZipPath).Items
    End With

    UnZip = True
    Exit Function

ErrorCatch:

    MsgBox "�w���Zip�t�@�C����������܂���B"

    UnZip = False

End Function

''
' �w��̃p�X���t�H���_�ł��邩�m�F����
'
' �y�Q�l�z
' ZIP�`���ň��k�E�𓀂��s��VBA�}�N�� | ���S�Ҕ��Y�^
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath �m�F����p�X
'
' @return {Boolean} �t�H���_���ǂ���
'
Private Function isFolder(ByVal srcPath As String) As Boolean
    isFolder = pFso.FolderExists(srcPath)
End Function

''
' �w��̃p�X���t�@�C���ł��邩�m�F����
'
' �y�Q�l�z
' ZIP�`���ň��k�E�𓀂��s��VBA�}�N�� | ���S�Ҕ��Y�^
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath �m�F����p�X
'
' @return {Boolean} �t�@�C�����ǂ���
'
Private Function isFile(ByVal srcPath As String) As Boolean
    isFile = pFso.FileExists(srcPath)
End Function

''
' ��؂蕶����}��
'
' �y�Q�l�z
' ZIP�`���ň��k�E�𓀂��s��VBA�}�N�� | ���S�Ҕ��Y�^
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath ���̃p�X
'
' @return {String} �}����̃p�X
'
Private Function addPathSeparator(ByVal srcPath As String) As String

    If Right(srcPath, 1) <> ChrW(92) Then srcPath = srcPath & ChrW(92)
    addPathSeparator = srcPath

End Function