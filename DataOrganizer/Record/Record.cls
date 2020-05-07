VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name            = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable       = False
Attribute VB_PredeclaredId   = False
Attribute VB_Exposed         = False
Attribute VB_Description     = "レコードを管理します。"

Option Explicit

Private pRecordId As Long
Private pStoreName As String
Private pEemployeeName As String
Private pPhoneNumber As String
Private pWage As Long

Property Get RecordId() As Long:                 RecordId = pRecordId: End Property
Property Let RecordId(ByVal aRecordId As Long): pRecordId = aRecordId: End Property

Property Get StoreName() As String:                  StoreName = pStoreName: End Property
Property Let StoreName(ByVal aStoreName As String): pStoreName = aStoreName: End Property

Property Get EmployeeName() As String:                     EmployeeName = pEmployeeName: End Property
Property Let EmployeeName(ByVal aEmployeeName As String): pEmployeeName = aEmployeeName: End Property

Property Get PhoneNumber() As String:                    PhoneNumber = pPhoneNumber: End Property
Property Let PhoneNumber(ByVal aPhoneNumber As String): pPhoneNumber = aPhoneNumber: End Property

Property Get Wage() As Long:             Wage = pWage: End Property
Property Let Wage(ByVal aWage As Long): pWage = aWage: End Property

Property Get Count() As Long
    Dim properties() As Variant
    properties = Array( _
        pRecordId, _
        pStoreName, _
        pEemployeeName, _
        pPhoneNumber, _
        pWage _
    )

    Count = UBound(properties) + 1

End Property

''
' コンストラクタ
'
Sub Init(ByVal aRecordId As Long)

    pRecordId = aRecordId
    pStoreName = ""
    pEemployeeName = ""
    pPhoneNumber = ""
    pWage = Empty

End Sub
