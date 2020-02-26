VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'������������������ ���W���[���ϐ� ��������������������

Private pApps As Application
Private pStartDate As Single

Private pTime As Double
Private pStartTime As Double
Private pFinishTime As Double

'������������������ �v���p�e�B ��������������������

Property Get Apps() As Application: Set Apps = pApps: End Property
Property Get StartDate() As Double: StartDate = pStartDate: End Property

Property Get Time() As Double
    With Me
        pTime = .FinishTime - .StartTime
    End With
    Time = pTime
End Property

Property Get StartTime() As Double: StartTime = pStartTime: End Property

Property Get FinishTime() As Double: FinishTime = pFinishTime: End Property

'���������������� �R���X�g���N�^�� ������������������

'����������
Private Sub Class_Initialize()

    Set pApps = Application
    pStartDate = Now

'    Debug.Print "The Instance for 'IpTool' is initialized."

End Sub

'�I��������
Private Sub Class_Terminate()

    ' ���̂��ƕϐ���j�����Ă��܂��̂ŁA��ɕ\������
'    Debug.Print "The Terminate for 'IpTool' is called."


End Sub

'�������������������� ���\�b�h ��������������������

''
' �X�^�[�g
'
Public Sub Start()
    pStartTime = Timer
End Sub

'''
'' �ꎞ��~
''
'Public Sub Pause()
'
'    '�쐬��
'
'End Sub

''
' �J�E���g���X�^�[�g
'
Public Sub Finish()
    pFinishTime = Timer
End Sub

''
' �������Ԃ��L�^
'
' @param {Worksheet} targetSheet   �L�^��
' @param {Long}      trialTimes    ���s��
' @param {String}    procedureName �v���V�[�W����
'
Public Sub Log(ByRef targetSheet As Worksheet, _
               Optional ByVal trialTimes As Long = 1, _
               Optional ByVal procedureName As String = "")
    
    If targetSheet Is Nothing Then
        MsgBox "�����ɃV�[�g���ݒ肳��Ă��܂���B"
        Stop
    End If
    
    With Me
        If .FinishTime = 0 Then .Finish
    
        Dim values() As Variant
        values = Array(.StartDate, .Time, trialTimes, procedureName)
        End With
        
        With targetSheet
            Dim lastRow As Long
            
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
            .Cells(lastRow, 1).Resize(1, UBound(values) + 1) = values
'            .Range(.Cells(lastRow, 1), .Cells(lastRow, UBound(values) + 1)).Value = values
    End With

End Sub