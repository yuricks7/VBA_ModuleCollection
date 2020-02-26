VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceImprovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'������������������ �v���p�e�B ��������������������

Private pApps As Application
Private pStartDate As Double

Private pTime As Double
Private pStartTime As Double
Private pFinishTime As Double

Property Get Apps() As Application: Set Apps = pApps: End Property
Property Get StartDate() As Double
    StartDate = pStartDate
End Property

Property Get Time() As Double
    With Me
        pTime = .FinishTime - .StartTime
    End With
    Time = pTime
End Property

Property Get StartTime() As Double:   StartTime = pStartTime:  End Property
Property Get FinishTime() As Double: FinishTime = pFinishTime: End Property


'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()

    Set pApps = Application
    pStartDate = Now()

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'�I��������
Private Sub Class_Terminate()

    ' ���̂��ƕϐ���j�����Ă��܂��̂ŁA��ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."

End Sub

'�������������������� ���\�b�h ��������������������
Public Sub Start()

    '��ʂ̍ĕ`��/�����v�Z/�C�x���g��t���~
    With Me.Apps
        .StatusBar = "��ƊJ�n���܂��[(�0�)�U"
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
        .Wait [Now()] + 50 / 86400000 '�`��ĊJ�Ɏ��Ԃ������邱�Ƃ����邽��
        .StatusBar = "���܂������܂����[(�0�)�U"
    End With

End Sub

Public Sub SetTimer()
    pStartTime = Timer
End Sub

Public Sub PauseTimer()



End Sub

Public Sub StopTimer()
    pFinishTime = Timer
End Sub


Public Sub Log(Optional ByRef targetSheet As Worksheet, _
               Optional ByVal trialTimes As Long = 1, _
               Optional ByVal procedureName As String = "")
    
    If targetSheet Is Nothing Then
        Set targetSheet = shProcLog
    End If

    With Me
        If .FinishTime = 0 Then .StopTimer
    
        Dim values() As Variant
        values = Array(.StartDate, .Time, trialTimes, procedureName)
        End With
        
        With targetSheet
            Dim lastRow As Long
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            .Range(.Cells(lastRow, 1), .Cells(lastRow, UBound(values) + 1)).Value = values
    End With
End Sub


Public Sub CursorMovesTo(ByRef startPosition As Range)

    '���O����
    '��ʕ`�悪�~�܂��Ă���ꍇ�A�J�[�\�������ړ�����
    '�E�B���h�E�ʒu�͂��̂܂܂ɂȂ��Ă��܂��̂ŁA
    '�O�̂��߉�ʕ`����ĊJ���Ă���
    Application.ScreenUpdating = True

    '�u�Z��.Select�v�̃G���[�h�~
    startPosition.Parent.Activate

    '�����C��
    With startPosition
        '�i�Ɨ����������������c�H�j
        If .Value = "" Then Set startPosition = .End(xlDown)
    End With

    '�đ�������s����Ă���������With�u���b�N�őI������Ă���startPosition��
    '�đ���O�̒l�̂܂܂Ȃ̂ŁAWith�̊O���ŉ��߂đI��
    startPosition.Select

    '���㏈��
    '���ɖ߂��Ă���
    Application.ScreenUpdating = False

End Sub


Public Sub ProgressBar(ByVal totalCounts As Long, _
                       ByVal currentCounts As Long)
    With Me
        '���݂̎��Ԃ��v�Z
        .StopTimer
        Dim remained As Double
        remained = .Time / currentCounts * (totalCounts - currentCounts)

        Apps.StatusBar = "�o�͒�... [��] " & Format(currentCounts, "#,##0") & " �� / " & Format(totalCounts, "#,##0") & " ��"
    End With

End Sub


Public Sub WaitingAnime(Optional ByVal currentCounts As Long = 300)

    With Me
        '�A�j��
        Dim dotMark As String
        Dim quotient As Double
        quotient = currentCounts / 3000
        Select Case quotient
            Case Is < 3
                dotMark = ""
            Case Is < 6
                dotMark = "."
            Case Is < 9
                dotMark = ".."
            Case Is >= 9
                dotMark = "..."
        End Select

        With .Apps
            .StatusBar = "�f�[�^����͂��Ă��܂�" & dotMark
        End With
    End With

End Sub