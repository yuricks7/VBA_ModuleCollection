VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���������������� ���W���[�����x���ϐ� ������������������

Private pApps As Application

'���������������� �v���p�e�B ������������������

Property Get Apps() As Application: Set Apps = pApps: End Property

'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()

    Set pApps = Application

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'���������������� ���\�b�h ������������������

Public Sub Start()

    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���~
    With Me.Apps
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

End Sub

Public Sub Finish()

    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���ĊJ
    With Me.Apps
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub