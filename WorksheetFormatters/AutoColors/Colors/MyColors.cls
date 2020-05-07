VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "MyColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "よく使う色のまとめ。"
Option Explicit

'列挙体
Public Enum eRGB
    Red
    Green
    Blue
    All
End Enum

'色の名前
Private pLightBlue As Long
Private pLighterBlue As Long
Private pNavyBlue As Long
Private pDarkBlue As Long

Private pDarkRed As Long
Private pPink As Long
Private pLightOrange As Long

Private pLightYellow As Long
Private pLightGreen As Long

Private pWhite As Long
Private pLightGray As Long
Private pDarkGray As Long
Private pBlack As Long

'色を定数としてセット
   Property Get LightBlue() As Long:    LightBlue = RGB(142, 169, 219): End Property
 Property Get LighterBlue() As Long:  LighterBlue = RGB(189, 215, 238): End Property
    Property Get DarkBlue() As Long:     DarkBlue = RGB(48, 84, 150): End Property
    Property Get NavyBlue() As Long:     NavyBlue = RGB(0, 32, 96): End Property '表のヘッダー向け

     Property Get DarkRed() As Long:      DarkRed = RGB(192, 0, 0): End Property
        Property Get Pink() As Long:         Pink = RGB(253, 173, 253): End Property
 Property Get LightOrange() As Long:  LightOrange = RGB(251, 226, 209): End Property

 Property Get LightYellow() As Long:  LightYellow = RGB(255, 242, 204): End Property
Property Get LighterGreen() As Long: LighterGreen = RGB(226, 239, 218): End Property '表のヘッダー向け
  Property Get LightGreen() As Long:   LightGreen = RGB(198, 224, 180): End Property

   Property Get LightGray() As Long:    LightGray = RGB(217, 217, 217): End Property
    Property Get DarkGray() As Long:     DarkGray = RGB(166, 166, 166): End Property
       Property Get Black() As Long:        Black = RGB(0, 0, 0): End Property
       Property Get White() As Long:        White = RGB(255, 255, 255): End Property

'━━━━━━━━ コンストラクタ等 ━━━━━━━━━

'初期化処理
Private Sub Class_Initialize()


End Sub


'終了時処理
Private Sub Class_Terminate()

    ' このあと変数を破棄してしまうので、先に表示する
'    Debug.Print "The Terminate for 'IpTool' is called."

End Sub

Public Function AddRgbComponent(ByVal colorValue As Long, _
                                Optional ByVal redValue As Long = 0, _
                                Optional ByVal greenValue As Long = 0, _
                                Optional ByVal blueValue As Long = 0) As Long

    Dim tempRgb As Variant
    tempRgb = GetRgbValues(colorValue, All)

    AddRgbComponent = RGB(tempRgb(0) + redComp, _
                          tempRgb(1) + greenComp, _
                          tempRgb(2) + blueComp)

End Function

Public Function AddRed(ByVal colorValue As Long) As Long

    Dim tempRgb As Variant
    tempRgb = GetRgbValues(colorValue, All)

    AddRed = RGB(tempRgb(0) + 5, tempRgb(1), tempRgb(2))

End Function

Public Function AddGreen(ByVal colorValue As Long) As Long

    Dim tempRgb As Variant
    tempRgb = GetRgbValues(colorValue, All)

    AddGreen = RGB(tempRgb(0), tempRgb(1) + 5, tempRgb(2))

End Function

Public Function AddBlue(ByVal colorValue As Long) As Long

    Dim tempRgb As Variant
    tempRgb = GetRgbValues(colorValue, All)

    AddBlue = RGB(tempRgb(0), tempRgb(1), tempRgb(2) + 5)

End Function


'Color値をRGB関数値に変換 | Excel作業をVBAで効率化
'https://vbabeginner.net/color%E5%80%A4%E3%82%92rgb%E9%96%A2%E6%95%B0%E5%80%A4%E3%81%AB%E5%A4%89%E6%8F%9B/

Public Function GetRgbValues(ByVal colorValue As Long, _
                             ByVal whichIsNeed As eRGB) As Variant

    Dim redValue As Long
    redValue = colorValue Mod 256

    Dim greenValue As Long
    greenValue = Int(colorValue / 256) Mod 256

    Dim blueValue As Long
    blueValue = Int(colorValue / 256 / 256)

    Dim ret As Variant
    Select Case whichIsNeed
        Case eRGB.Red
            ret = redValue

        Case eRGB.Green
            ret = greenValue

        Case eRGB.Blue
            ret = blueValue

        Case Else
            ret = Array(redValue, greenValue, blueValue)

    End Select

    GetRgbValues = ret

End Function

Public Function GetHexString(ByVal colorValue As Long) As String
    '6桁すべて数値の場合も考えて念のため文字列化
    GetHexValue = CStr(Hex(colorValue))
End Function
