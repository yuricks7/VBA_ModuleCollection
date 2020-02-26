Attribute VB_Name = "fomatNumbers"
Option Explicit

Sub test()

    Dim nFormat As numberFormatter
    Set nFormat = New numberFormatter
    
    Debug.Print nFormat.NumberFormat("0,000", "Й~", Positive_Negative_Zero, BlueBlack, "Бе", RedPurple, LightGreen)
    
End Sub
