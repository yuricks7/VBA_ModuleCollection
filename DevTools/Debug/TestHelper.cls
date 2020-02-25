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
'** 【編集中】                               **
'** シフト作成ツールから抜粋                 **
'**                                          **
'** 【参照設定】                             **
'** [Name] ●●●●                          **
'** [note] ●●●●●●●●●●●●          **
'**********************************************

'━━━━━━━━ モジュールレベル変数 ━━━━━━━━━

''モジュールレベル変数
Private pFolderPathArray As Long
'
''━━━━━━━━ プロパティ ━━━━━━━━━
'
Public Property Get FolderPathArray() As Variant: FolderPathArray = pFolderPathArray: End Property
Public Property Let FolderPathArray(ByVal aFolderPathArray As Variant): pFolderPathArray = aFolderPathArray: End Property
'
''━━━━━━━━ コンストラクタ等 ━━━━━━━━━
'
''初期化処理
'Private Sub Class_Initialize()
'
'
'End Sub
'
'
''終了時処理
'Private Sub Class_Terminate()
''    使うときは先に表示する
''    Debug.Print "The Terminate for 'IpTool' is called."
'
'
'End Sub

'━━━━━━━━ メソッド ━━━━━━━━━

''
' マクロ実行前の状態にする
'
Public Sub Reset(Optional ByRef folderPaths As Variant)

    Call CloseBooks
        
    If IsEmpty(folderPaths) Then folderPaths = Me.FolderPathArray
    Call DeleteOldFiles(folderPaths)
    
End Sub

''
' このブック以外をまとめて閉じる
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
    
    MsgBox "このマクロ搭載ブック以外、保存せずに閉じました。"

End Sub

''
' 作成済みフォルダーを空にする
'
' 【参考】
' フォルダ（サブフォルダも全て）削除する、Optionでファイルのみ削除 | Excelの神髄
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' @param {Variant} folderPaths 削除を実行するフォルダのパス（配列）
'
Public Sub DeleteOldFiles(ByRef folderPaths As Variant)

    Dim i As Long
    For i = 0 To UBound(folderPaths)
        Dim resultMsg As String
        
        Dim hasNoError As Boolean
        hasNoError = deleteAllFilesIn(folderPaths(i), resultMsg, True)
        
        If hasNoError Then
            MsgBox "削除完了。"
        
        Else
            MsgBox resultMsg
        
        End If
    Next

End Sub

''
' フォルダ内を空にする
'
' 【参考】
' フォルダ（サブフォルダも全て）削除する、Optionでファイルのみ削除 | Excelの神髄
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' 【参照設定】
' [Name] Microsoft Scripting Runtime
' [note] 参照設定しない場合
'   Dim objFSO As Object
'   Set objFSO = CreateObject("Scripting.FileSystemObject")
'
' @param {String}  targetDirPath 空にするフォルダ
' @param {String}  resultMsg     実行結果
' @param {Boolean} isOnlyFile    ファイルだけ削除ならTrue
'
' @return {Boolean} 完了/失敗
'
Private Function deleteAllFilesIn(ByVal targetDirPath As String, _
                                  ByRef resultMsg As String, _
                                  Optional ByVal isOnlyFile As Boolean = False) As Boolean

'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    resultMsg = ""
    
    'フォルダの存在チェック
    If Not fso.FolderExists(targetDirPath) Then
        resultMsg = "指定のフォルダは存在しません。"
        deleteAllFilesIn = False
        Exit Function
    End If

    '削除を実行
    Dim folderObj As Folder
    Set folderObj = fso.GetFolder(targetDirPath)
    Call deleteDirectories(folderObj, isOnlyFile, resultMsg)

    'エラーの有無を確認
    If resultMsg = "" Then
        deleteAllFilesIn = True
    
    Else
        deleteAllFilesIn = False
    
    End If

End Function

''
' 削除の実行
'
' 【参考】
' フォルダ（サブフォルダも全て）削除する、Optionでファイルのみ削除 | Excelの神髄
' https://excel-ubara.com/excelvba5/EXCELVBA259.html
'
' @param {Folder}  folderObj  fsoオブジェクト
' @param {Boolean} isOnlyFile ファイルだけ削除ならTrue
' @param {String}  resultMsg  実行結果
'
Private Sub deleteDirectories(ByVal folderObj As Folder, _
                              ByVal isOnlyFile As Boolean, _
                              ByRef resultMsg As String)

    Dim subFolderObj As Folder
    Dim fileObj As File
    
    On Error Resume Next
    
    'サブフォルダを探索
    For Each subFolderObj In folderObj.SubFolders
        Call deleteDirectories(subFolderObj, isOnlyFile, resultMsg) '再帰
    Next

    Dim errMsg As String
    
    'ファイルを削除
    For Each fileObj In folderObj.Files
        fileObj.Delete
        
        If Err.Number <> 0 Then
            errMsg = "ファイル「" & fileObj.Path & "」が削除できませんでした"
            resultMsg = resultMsg & errMsg & vbLf
            
            Err.Clear
        End If
    Next

    'フォルダを削除
    If isOnlyFile Then GoTo FinishDelete
    
    folderObj.Delete
    If Err.Number <> 0 Then
        errMsg = "フォルダ「" & folderObj.Path & "」が削除できませんでした"
        resultMsg = resultMsg & errMsg & vbLf
        
        Err.Clear
    End If
        
FinishDelete:
    Set subFolderObj = Nothing
    Set fileObj = Nothing
        
    On Error GoTo 0

End Sub

