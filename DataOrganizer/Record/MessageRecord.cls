VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRecordId As Long
Private pPostTime As Date
Private pUserId As String
Private pUserName As String
Private pLines As New PowerArray
Private pBody As String
Private pHasImage As Boolean

Property Get RecordId() As Long:                 RecordId = pRecordId: End Property
Property Let RecordId(ByVal aRecordId As Long): pRecordId = aRecordId: End Property

Property Get PostTime() As Date:                 PostTime = pPostTime: End Property
Property Let PostTime(ByVal aPostTime As Date): pPostTime = aPostTime: End Property

Property Get UserId() As String:               UserId = pUserId: End Property
Property Let UserId(ByVal aUserId As String): pUserId = aUserId: End Property

Property Get UserName() As String:                 UserName = pUserName: End Property
Property Let UserName(ByVal aUserName As String): pUserName = aUserName: End Property

Property Get Lines() As PowerArray:              Set Lines = pLines: End Property
Property Set Lines(ByVal aLines As PowerArray): Set pLines = aLines: End Property

Property Get Body() As String
    pBody = Me.Lines.JoinVia(vbLf)

    Body = pBody
End Property

Property Let Body(ByVal aBody As String)
    pBody = aBody
End Property

Property Get HasImage() As Boolean:                 HasImage = pHasImage: End Property
Property Let HasImage(ByVal aHasImage As Boolean): pHasImage = aHasImage: End Property

Property Get Count() As Long
    Count = 6

'    以下6つ
'    Private pRecordId As Long
'    Private pPostTime As Date
'    Private pUserId As String
'    Private pUserName As String
'    Private pLines As New PowerArray
'    Private pHasImage As Boolean

End Property

''
' コンストラクタ
'
Sub Init(ByVal aRecordId As Long)

    pRecordId = aRecordId

    pPostTime = Empty
    pUserId = ""
    pUserName = ""

    Set pLines = New PowerArray
    pLines.Init

    pBody = ""

End Sub
