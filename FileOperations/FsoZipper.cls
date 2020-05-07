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
' ファイル・フォルダをZIP形式で圧縮
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {Variant} srcPath        元ファイル・フォルダ
' @param {Variant} dstFolderPath  出力先（指定しない場合は元ファイル・フォルダと同じ場所）
'
' @return {Boolean} 圧縮の成否
'
Public Function Zip(ByVal srcPath As Variant, _
                    Optional ByVal dstFolderPath As Variant = "") As Boolean

    On Error GoTo ErrorCatch

    With pFso
        'パスの種類をチェック
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
        '出力先の設定
        Dim dstFilePath As Variant
        dstFilePath = addPathSeparator(dstFolderPath) & .GetBaseName(srcPath) & ".zip"

        '空のZIPファイルを作成
        With .CreateTextFile(dstFilePath, True)
          .Write ChrW(&H50) & ChrW(&H4B) & ChrW(&H5) & ChrW(&H6) & String(18, ChrW(0))
          .Close
        End With

    End With

    'ZIPファイルに元ファイルのコピーを格納
    '（フォルダもフォルダ毎コピーされる）
    With CreateObject("Shell.Application")
        With .Namespace(dstFilePath)
'効かない
'            Application.DisplayAlerts = False
            .CopyHere srcPath
'            Application.DisplayAlerts = True

            Do While .Items.Count < 1
                DoEvents 'システムに制御を渡す
            Loop
        End With
    End With

    Zip = True
    Exit Function

ErrorCatch:

    MsgBox "圧縮を遂行できませんでした。"

    Zip = False

End Function

''
' ZIPファイルを解凍
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {Variant} srcZipPath    Zipファイルのパス
' @param {Variant} dstFolderPath 展開先（指定しない場合は元ファイル・フォルダと同じ場所）
'                                ※展開先に同名ファイルがあった場合はユーザー判断で処理
'
' @return {Boolean} 解凍の成否
'
Public Function UnZip(ByVal srcZipPath As Variant, _
                      Optional ByVal dstFolderPath As Variant = "") As Boolean

    '実行前チェック
    With pFso
        If .FileExists(srcZipPath) = False Then GoTo ErrorCatch 'ファイルの存在確認
        If LCase(.GetExtensionName(srcZipPath)) <> "zip" Then GoTo ErrorCatch '拡張子

        'ファイルを指定している場合は、その親フォルダに変更
        If isFolder(dstFolderPath) = False Then
            dstFolderPath = .GetFile(srcZipPath).ParentFolder.Path
        End If
    End With

    '解凍
    With CreateObject("Shell.Application")
        .Namespace(dstFolderPath).CopyHere .Namespace(srcZipPath).Items
    End With

    UnZip = True
    Exit Function

ErrorCatch:

    MsgBox "指定のZipファイルが見つかりません。"

    UnZip = False

End Function

''
' 指定のパスがフォルダであるか確認する
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath 確認するパス
'
' @return {Boolean} フォルダかどうか
'
Private Function isFolder(ByVal srcPath As String) As Boolean
    isFolder = pFso.FolderExists(srcPath)
End Function

''
' 指定のパスがファイルであるか確認する
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath 確認するパス
'
' @return {Boolean} ファイルかどうか
'
Private Function isFile(ByVal srcPath As String) As Boolean
    isFile = pFso.FileExists(srcPath)
End Function

''
' 区切り文字を挿入
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {String} SrcPath 元のパス
'
' @return {String} 挿入後のパス
'
Private Function addPathSeparator(ByVal srcPath As String) As String

    If Right(srcPath, 1) <> ChrW(92) Then srcPath = srcPath & ChrW(92)
    addPathSeparator = srcPath

End Function
