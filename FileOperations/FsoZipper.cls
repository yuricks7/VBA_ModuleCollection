VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FsoZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' @param {Variant} SrcPath        元ファイル・フォルダ
' @param {Variant} DestFolderPath 出力先（指定しない場合は元ファイル・フォルダと同じ場所）
'
Public Sub Zip(ByVal SrcPath As Variant, _
               Optional ByVal DestFolderPath As Variant = "")
    
    With pFso
        'パスの種類をチェック
        If isFolder(DestFolderPath) = False Then
            If isFolder(SrcPath) = True Then
                DestFolderPath = SrcPath
            
            ElseIf isFile(SrcPath) = True Then
                DestFolderPath = .GetFile(SrcPath).ParentFolder.Path
            
            Else
                Exit Sub
            
            End If
        End If
        
        '出力先の設定
        Dim DestFilePath As Variant
        DestFilePath = addPathSeparator(DestFolderPath) & .GetBaseName(SrcPath) & ".zip"
        
        '空のZIPファイルを作成
        With .CreateTextFile(DestFilePath, True)
          .Write ChrW(&H50) & ChrW(&H4B) & ChrW(&H5) & ChrW(&H6) & String(18, ChrW(0))
          .Close
        End With
    
    End With
     
    'ZIPファイルにコピーを格納
    With CreateObject("Shell.Application")
        With .Namespace(DestFilePath)
            .CopyHere SrcPath
            
            While .Items.Count < 1
                DoEvents
            Wend
        End With
    End With

End Sub

''
' ZIPファイルを解凍
'
' 【参考】
' ZIP形式で圧縮・解凍を行うVBAマクロ | 初心者備忘録
' https://www.ka-net.org/blog/?p=7605
'
' @param {Variant} SrcPath        元ファイル
' @param {Variant} DestFolderPath 出力先（指定しない場合は元ファイル・フォルダと同じ場所）
'                                 ※出力先に同名ファイルがあった場合はユーザー判断で処理
'
Public Sub UnZip(ByVal SrcPath As Variant, _
                 Optional ByVal DestFolderPath As Variant = "")
  
    '実行前チェック
    With pFso
        If .FileExists(SrcPath) = False Then Exit Sub               '存在確認
        If LCase(.GetExtensionName(SrcPath)) <> "zip" Then Exit Sub '拡張子
        
        'ファイルを指定している場合は、その親フォルダに変更
        If isFolder(DestFolderPath) = False Then
            DestFolderPath = .GetFile(SrcPath).ParentFolder.Path
        End If
    End With
    
    '解凍
    With CreateObject("Shell.Application")
        .Namespace(DestFolderPath).CopyHere .Namespace(SrcPath).Items
    End With
End Sub
 
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
Private Function isFolder(ByVal SrcPath As String) As Boolean
    isFolder = pFso.FolderExists(SrcPath)
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
Private Function isFile(ByVal SrcPath As String) As Boolean
    isFile = pFso.FileExists(SrcPath)
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
Private Function addPathSeparator(ByVal SrcPath As String) As String
  
    If Right(SrcPath, 1) <> ChrW(92) Then SrcPath = SrcPath & ChrW(92)
    addPathSeparator = SrcPath

End Function
