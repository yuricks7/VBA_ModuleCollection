VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************************
'** �y�ҏW���z                               **
'** �V�t�g�쐬�c�[�����甲��                 **
'**                                          **
'** �y�Q�Ɛݒ�z                             **
'** [Name] ��������                          **
'** [note] ������������������������          **
'**********************************************

'���������������� ���W���[�����x���ϐ� ������������������

''���W���[�����x���ϐ�
Private pFolderPathArray As Long
'
''���������������� �v���p�e�B ������������������
'
Public Property Get FolderPathArray() As Variant: FolderPathArray = pFolderPathArray: End Property
Public Property Let FolderPathArray(ByVal aFolderPathArray As Variant): pFolderPathArray = aFolderPathArray: End Property
'
''���������������� �R���X�g���N�^�� ������������������
'
''����������
'Private Sub Class_Initialize()
'
'
'End Sub
'
'
''�I��������
'Private Sub Class_Terminate()
''    �g���Ƃ��͐�ɕ\������
''    Debug.Print "The Terminate for 'IpTool' is called."
'
'
'End Sub

'���������������� ���\�b�h ������������������

''
' �}�N�����s�O�̏�Ԃɂ���
'
Public Sub Reset(Optional ByRef folderPaths As Variant)

    Call CloseBooks
        
    If IsEmpty(folderPaths) Then folderPaths = Me.FolderPathArray
    Call DeleteOldFiles(folderPaths)
    
End Sub

''
' ���̃u�b�N�ȊO���܂Ƃ߂ĕ���
'
Public Sub CloseBooks()

    Dim wb As Workbook
    For Each wb In Workbooks
        With wb
            If .Name <> ThisWorkbook.Name Then
                .Close SaveChanges:=False
            End If
        End With
    Next
    
    MsgBox "���̃}�N�����ڃu�b�N�ȊO�A�ۑ������ɕ��܂����B"

End Sub

''
' �쐬�ς݃t�H���_�[����ɂ���
'
' �y�Q�l�z
' �t�H���_�i�T�u�t�H���_���S�āj�폜����AOption�Ńt�@�C���̂ݍ폜 | Excel�̐_��
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' @param {Variant} folderPaths �폜�����s����t�H���_�̃p�X�i�z��j
'
Public Sub DeleteOldFiles(ByRef folderPaths As Variant)

    Dim i As Long
    For i = 0 To UBound(folderPaths)
        Dim resultMsg As String
        
        Dim hasNoError As Boolean
        hasNoError = deleteAllFilesIn(folderPaths(i), resultMsg, True)
        
        If hasNoError Then
            MsgBox "�폜�����B"
        
        Else
            MsgBox resultMsg
        
        End If
    Next

End Sub

''
' �t�H���_������ɂ���
'
' �y�Q�l�z
' �t�H���_�i�T�u�t�H���_���S�āj�폜����AOption�Ńt�@�C���̂ݍ폜 | Excel�̐_��
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' �y�Q�Ɛݒ�z
' [Name] Microsoft Scripting Runtime
' [note] �Q�Ɛݒ肵�Ȃ��ꍇ
'   Dim objFSO As Object
'   Set objFSO = CreateObject("Scripting.FileSystemObject")
'
' @param {String}  targetDirPath ��ɂ���t�H���_
' @param {String}  resultMsg     ���s����
' @param {Boolean} isOnlyFile    �t�@�C�������폜�Ȃ�True
'
' @return {Boolean} ����/���s
'
Private Function deleteAllFilesIn(ByVal targetDirPath As String, _
                                  ByRef resultMsg As String, _
                                  Optional ByVal isOnlyFile As Boolean = False) As Boolean

'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    resultMsg = ""
    
    '�t�H���_�̑��݃`�F�b�N
    If Not fso.FolderExists(targetDirPath) Then
        resultMsg = "�w��̃t�H���_�͑��݂��܂���B"
        deleteAllFilesIn = False
        Exit Function
    End If

    '�폜�����s
    Dim folderObj As Folder
    Set folderObj = fso.GetFolder(targetDirPath)
    Call deleteDirectories(folderObj, isOnlyFile, resultMsg)

    '�G���[�̗L�����m�F
    If resultMsg = "" Then
        deleteAllFilesIn = True
    
    Else
        deleteAllFilesIn = False
    
    End If

End Function

''
' �폜�̎��s
'
' �y�Q�l�z
' �t�H���_�i�T�u�t�H���_���S�āj�폜����AOption�Ńt�@�C���̂ݍ폜 | Excel�̐_��
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' @param {Folder}  folderObj  fso�I�u�W�F�N�g
' @param {Boolean} isOnlyFile �t�@�C�������폜�Ȃ�True
' @param {String}  resultMsg  ���s����
'
Private Sub deleteDirectories(ByVal folderObj As Folder, _
                              ByVal isOnlyFile As Boolean, _
                              ByRef resultMsg As String)

    Dim subFolderObj As Folder
    Dim fileObj As File
    
    On Error Resume Next
    
    '�T�u�t�H���_��T��
    For Each subFolderObj In folderObj.SubFolders
        Call deleteDirectories(subFolderObj, isOnlyFile, resultMsg) '�ċA
    Next

    Dim errMsg As String
    
    '�t�@�C�����폜
    For Each fileObj In folderObj.Files
        fileObj.Delete
        
        If Err.Number <> 0 Then
            errMsg = "�t�@�C���u" & fileObj.Path & "�v���폜�ł��܂���ł���"
            resultMsg = resultMsg & errMsg & vbLf
            
            Err.Clear
        End If
    Next

    '�t�H���_���폜
    If isOnlyFile Then GoTo FinishDelete
    
    folderObj.Delete
    If Err.Number <> 0 Then
        errMsg = "�t�H���_�u" & folderObj.Path & "�v���폜�ł��܂���ł���"
        resultMsg = resultMsg & errMsg & vbLf
        
        Err.Clear
    End If
        
FinishDelete:
    Set subFolderObj = Nothing
    Set fileObj = Nothing
        
    On Error GoTo 0

End Sub
