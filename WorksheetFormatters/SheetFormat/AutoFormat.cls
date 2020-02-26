VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���������������������� ���W���[���ϐ� ����������������������

Private pFont As String
Private pFontSize As Long
Private pHeaderFontSize As Long

Private pVerticalAlignment As Long
Private pRowHeight As Double

Private pHorizontalAlignment As Long
Private pColumnWidth_MAX As Double
Private pColumnWidth_MIDDLE As Double
Private pColumnWidth_NARROW As Double

Private pColumnWidth_GridPaper As Double
Private pColumnWidth_MinutesAndHour As Double
Private pColumnWidth_Date_WithDayOfTheWeek As Double
Private pColumnWidth_DateTime_WithDayOfTheWeek As Double
Private pColumnWidth_Boolean As Double

Private pTypeOfStandard As String
Private pTypeOfString As String
Private pTypeOfNumbers As String
Private pTypeOfNumbers00 As String

Private pTypeOfDate As String
Private pTypeOfDateTime As String
Private pTypeOfDate_WithDayOfTheWeek As String
Private pTypeOfDateTime_WithDayOfTheWeek As String

'���������������������� �v���p�e�B ����������������������

''
' �t�H���g
'
Property Get Font() As String
    
    pFont = "Meiryo UI"
    Font = pFont

End Property

Property Get FontSize() As Long
    
    pFontSize = 8
    FontSize = pFontSize

End Property

Property Get HeaderFontSize() As Long
    
    pHeaderFontSize = 9
    HeaderFontSize = pHeaderFontSize

End Property

''
' �W���̐����ʒu
'
Property Get VerticalAlignment() As Long
    
    pVerticalAlignment = xlBottom
    VerticalAlignment = pVerticalAlignment

End Property

Property Let VerticalAlignment(ByVal aVerticalAlignment As Long)

    '����l�̎g���������Ă�c�H
    If IsEmpty(aVerticalAlignment) Then aVerticalAlignment = xlBottom
    pVerticalAlignment = aVerticalAlignment

End Property

''
' �s�̍���
'
Property Get RowHeight() As Double
    
    pRowHeight = 14.4 '24px
    RowHeight = pRowHeight

End Property

''
' �W���̐����ʒu
'
Property Get HorizontalAlignment() As Long
    
    pHorizontalAlignment = xlRight
    HorizontalAlignment = pHorizontalAlignment

End Property

Property Let HorizontalAlignment(ByVal aHorizontalAlignment As Long)

    '����l�̎g���������Ă�c�H
    If IsEmpty(aHorizontalAlignment) Then aHorizontalAlignment = xlBottom
    pHorizontalAlignment = aHorizontalAlignment

End Property

''
' ��
'
Property Get ColumnWidth_MAX() As Double
    
    pColumnWidth_MAX = 46.88 '380px
    ColumnWidth_MAX = pColumnWidth_MAX

End Property

Property Get ColumnWidth_MIDDLE() As Double
    
    pColumnWidth_MIDDLE = 34.38 '280px
    ColumnWidth_MIDDLE = pColumnWidth_MIDDLE

End Property

Property Get ColumnWidth_NARROW() As Double
    
    pColumnWidth_NARROW = 14.38 '120px
    ColumnWidth_NARROW = pColumnWidth_NARROW

End Property

'Excel���ᎆ
Property Get ColumnWidth_GridPaper() As Double
    
    pColumnWidth_GridPaper = 2.88 '28px
    ColumnWidth_GridPaper = pColumnWidth_GridPaper

End Property

'H��AM��
Property Get ColumnWidth_MinutesAndHour() As Double
    
    pColumnWidth_MinutesAndHour = 3.88 '36px
    ColumnWidth_MinutesAndHour = pColumnWidth_MinutesAndHour

End Property

Property Get ColumnWidth_Date_WithDayOfTheWeek() As Double
    
    pColumnWidth_Date_WithDayOfTheWeek = 14.38 '120px
    ColumnWidth_Date_WithDayOfTheWeek = pColumnWidth_Date_WithDayOfTheWeek

End Property

Property Get ColumnWidth_DateTime_WithDayOfTheWeek() As Double
    
    pColumnWidth_DateTime_WithDayOfTheWeek = 20.38 '168px
    ColumnWidth_DateTime_WithDayOfTheWeek = pColumnWidth_DateTime_WithDayOfTheWeek

End Property

'True / False
Property Get ColumnWidth_Boolean() As Double
    
    pColumnWidth_Boolean = 6.88 '60px
    ColumnWidth_Boolean = pColumnWidth_Boolean

End Property

''
' �Z���̕\���`��
'
Property Get TypeOfStandard() As String
    
    pTypeOfStandard = "G/�W��"
    TypeOfStandard = pTypeOfStandard

End Property

Property Get TypeOfString() As String
    
    pTypeOfString = "@"
    TypeOfString = pTypeOfString

End Property

Property Get TypeOfNumbers() As String
    
    pTypeOfNumbers = "#,##0_ ;[��]�� #,##0 "
    TypeOfNumbers = pTypeOfNumbers

End Property

Property Get TypeOfNumbers00() As String
    
    pTypeOfNumbers00 = "00"
    TypeOfNumbers00 = pTypeOfNumbers00

End Property

Property Get TypeOfDate() As String
    
    pTypeOfDate = "yyyy/mm/dd"
    TypeOfDate = pTypeOfDate

End Property

Property Get TypeOfDateTime() As String
    
    pTypeOfDateTime = "yyyy/mm/dd hh:mm"
    TypeOfDateTime = pTypeOfDateTime

End Property

Property Get TypeOfDate_WithDayOfTheWeek() As String

    pTypeOfDate_WithDayOfTheWeek = "yyyy/mm/dd (aaa)"
    TypeOfDate_WithDayOfTheWeek = pTypeOfDate_WithDayOfTheWeek

End Property

Property Get TypeOfDateTime_WithDayOfTheWeek() As String

    pTypeOfDateTime_WithDayOfTheWeek = "yyyy/mm/dd (aaa) hh:mm"
    TypeOfDateTime_WithDayOfTheWeek = pTypeOfDateTime_WithDayOfTheWeek

End Property

'���������������������� ���\�b�h ����������������������

''
' Excel�̕W���t�H���g��ς���
'
' �y�Q�l�z
' �yExcel VBA�z�u�b�N�̃t�H���g�A�܂Ƃ߂ĕύX���܂��񂩁H | ��邨���m�[�g
' https://www.yuru-wota.com/entry/VBA/set-favorite-font
'
Public Sub SetExcelFont(ByRef targetBook As Workbook)

    With Me
        Dim strFontName As String: strFontName = .Font
        Dim numFontSize As Long:   numFontSize = .FontSize
    End With

    'Excel�A�v���P�[�V�������̂̃t�H���g��ύX
    With Application
        .StandardFont = strFontName
        .StandardFontSize = numFontSize
    End With

    Dim ws As Worksheet
    For Each ws In targetBook.Worksheets
        With ws
            '�V�[�g�̕W���t�H���g�𓝈�
            With .Cells.Font
                    .Name = strFontName
                    .Size = numFontSize
            End With
            
            '�����̔z�u�𒲐�
            .Rows.RowHeight = Me.RowHeight
            .Cells.VerticalAlignment = Me.VerticalAlignment
        End With
    Next ws

    Dim msg As String
    Dim strFontSize As String: strFontSize = CStr(numFontSize)

    msg = "�ݒ��ύX���܂����B"
    msg = msg & vbCrLf & "----------------------------"
    msg = msg & vbCrLf & "��Excel�A�v���P�[�V����"
    msg = msg & vbCrLf & "�t�H���g�F" & strFontName
    msg = msg & vbCrLf & "�T�C�Y�F" & strFontSize & " px"
    msg = msg & vbCrLf & "�i�_�E�����[�h�����t�@�C���ɂ͔��f����܂���B�j"
    msg = msg & vbCrLf & "----------------------------"
    msg = msg & vbCrLf & "�����[�N�V�[�g"
    msg = msg & vbCrLf & "�t�H���g�F" & strFontName
    msg = msg & vbCrLf & "----------------------------"

    MsgBox msg

    MsgBox "�A�v���P�[�V�����̐ݒ��ύX�����ꍇ�A" & vbCrLf & _
                "�����ɔ��f���K�v�Ȏ��͈�UExcel���ċN�����Ă��������B"

End Sub

''
' �e�[�u���ƃV�[�g�̐ݒ���s��
'
Public Sub Apply(ByRef targetSheet As Worksheet)
    
    Dim targetBook As Workbook: Set targetBook = targetSheet.Parent
    If targetSheet Is Nothing Then Set targetSheet = targetBook.Worksheets(1)

    Const NEW_STYLE_NAME As String = "lightBlueBackStyle"

    With targetSheet
        Dim targetList As ListObject
        On Error GoTo NoTable
            Set targetList = .ListObjects(1)
        On Error GoTo 0
        
        '�X�^�C����ύX�i2��ڂ̓G���[�ɂȂ�悤�ȁH�j
        On Error GoTo CreateStyle
        targetList.TableStyle = NEW_STYLE_NAME
        On Error GoTo 0
        
        '�u�b�N�́u����̃X�^�C���v�ɐݒ�
        targetBook.DefaultTableStyle = NEW_STYLE_NAME
        
        '�N�G���̐ݒ��ύX
        If .QueryTables.Count > 0 Then
            Call SetQueryProperties(targetList)
        End If
        
        '�V�[�g�����e�[�u�����Ɠ���
        On Error Resume Next '�G���[�Ȃ�f�t�H���g�̂܂�
        targetSheet.Name = targetList.Name
        On Error GoTo 0
    End With
    
    
    '����n��
    Call SetSheetStyle(targetSheet)
    
    MsgBox "�e�[�u���𐮌`���܂����B", vbInformation
    
    Exit Sub

NoTable:
    MsgBox "�I�𒆂̃V�[�g�Ƀe�[�u��������܂���B"
    MsgBox "�����𒆎~���܂�� "
    Exit Sub

CreateStyle:
        Call createTableStyle(NEW_STYLE_NAME, targetBook)
        Resume '�G���[�����������X�e�[�g�����g�ɖ߂�

End Sub

''
' �e�[�u���̃X�^�C����ǉ�����
'
' @param {String}   newStyleName
' @param {Workbook} targetBook
'
Private Sub createTableStyle(ByVal newStyleName As String, _
                             ByRef targetBook As Workbook)
        
    '�X�^�C���̒ǉ�
    Dim newStyle As TableStyle
    Set newStyle = targetBook.TableStyles.Add(newStyleName)
    
    Call changeTableDesign(newStyle)
        
End Sub

''
' �e�[�u���X�^�C���̐ݒ��ύX����
'
' �y�Q�l�z
' �yExcel VBA�z�e�[�u���ϊ��ƃX�^�C���ύX - �e�[�u���ϊ�����Ƃ���܂ŕW�����W���[���ŏ����Ă݂� | ��邨���m�[�g
' https://www.yuru-wota.com/entry/VBA/set-table-safely
'
' @param {TableStyle} tableStyleObj �ݒ�Ɏg�p����TableStyle�I�u�W�F�N�g
'
Private Sub changeTableDesign(ByRef tableStyleObj As TableStyle)
    
    '�e�[�u���S�́iWholeStyle�j
    Dim black As Long:         black = RGB(0, 0, 0)
    Dim lightGray As Long: lightGray = RGB(208, 206, 206)
    
    Call setWholeStyle(tableStyleObj, black, lightGray)
    
    '���o���s�iHeaderRowStyle�j
    Dim deepBlue As Long: deepBlue = RGB(0, 32, 96)
    Dim white As Long:       white = RGB(255, 255, 255)

    Call setHeaderStyle(tableStyleObj, deepBlue, white, True)

End Sub

''
' �e�[�u���S�̂̐ݒ�
'
' @param {TableStyle} tableStyleObj  �ݒ�Ɏg�p����TableStyle�I�u�W�F�N�g
' @param {Long}       outerLineColor �O�g�̐F
' @param {Long}       innerLineColor �O�g�̐F
'
Private Sub setWholeStyle(ByRef tableStyleObj As Variant, _
                          Optional ByVal outerLineColor As Long = 0, _
                          Optional ByVal innerLineColor As Long = 13553360)

    Dim wholeTableElements As Variant
    Set wholeTableElements = tableStyleObj.TableStyleElements(xlWholeTable)
    
    Dim outerLineConstants As Variant
    outerLineConstants = Array( _
                             xlEdgeTop, _
                             xlEdgeBottom, _
                             xlEdgeLeft, _
                             xlEdgeRight _
                         )
    
    Call setLines( _
        wholeTableElements, _
        outerLineConstants, _
        outerLineColor, _
        xlContinuous, _
        xlMedium _
    )
    
    Dim innerLineConstants As Variant
    innerLineConstants = Array( _
                             xlInsideVertical, _
                             xlInsideHorizontal _
                         )
    
    Call setLines( _
        wholeTableElements, _
        innerLineConstants, _
        innerLineColor, _
        xlContinuous, _
        xlThin _
    )

End Sub

''
' �r����ݒ肷��
'
' @param {Variant} StyleElements �ݒ肷��StyleElements�I�u�W�F�N�g
' @param {Variant} linePositions ���̏ꏊ��\���g�ݍ��ݒ萔�̔z��
' @param {Long} targetColor      ���̐F
' @param {Long} lineStyle        [Optional] ���̎��
' @param {Long} thickness        [Optional] ���̑���
'
Private Sub setLines(ByRef StyleElements As Variant, _
                     ByRef linePositions As Variant, _
                     ByVal targetColor As Long, _
                     Optional ByVal lineStyle As Long = xlContinuous, _
                     Optional ByVal thickness As Long = xlThin)
 
    With StyleElements
        Dim i As Long
        For i = 0 To UBound(linePositions)
            With .Borders(linePositions(i))
                .Color = targetColor
                .lineStyle = lineStyle
                .Weight = thickness
            End With
        Next i
    End With

End Sub

''
' �w�b�_�[�̐ݒ�
'
' @param {TableStyle} tableStyleObj �ݒ肷��TableStyle�I�u�W�F�N�g
' @param {Long} interiorColor       [Optional] ���e
' @param {Long} fontColor           [Optional] ���e
' @param {Long} isBold              [Optional] ���e
'
Private Sub setHeaderStyle(ByRef tableStyleObj As TableStyle, _
                           Optional ByVal interiorColor As Long = 16777215, _
                           Optional ByVal fontColor As Long = 6299648, _
                           Optional ByVal isBold As Boolean = True)

   Dim headerRowElements As Variant
   Set headerRowElements = tableStyleObj.TableStyleElements(xlHeaderRow)
   
    With headerRowElements
        .Interior.Color = interiorColor
        
        With .Font
            .Color = fontColor
            .Bold = isBold
        End With
    
    End With

End Sub

''
' �N�G���e�[�u���̐ݒ�
'
' @param {ListObject} targetList �Ώۂ̃e�[�u��
'
Public Sub SetQueryProperties(ByRef targetList As ListObject)
    
    With targetList.QueryTable
        .RowNumbers = False           '�s�ԍ�
        .FillAdjacentFormulas = False '�N�G�� �e�[�u���̉E���̐����������I�ɍX�V
        .PreserveFormatting = True    '������ێ�����
        .RefreshOnFileOpen = False    '�t�@�C�����J������X�V����
        .BackgroundQuery = True       '�o�b�N�O���E���h�ōX�V�̋���
        .RefreshStyle = xlInsertDeleteCells '�V�������R�[�h�Z�b�g�ƍs���𐳊m�Ɉ�v������
        .SavePassword = False         'ODBC�ڑ�������̃p�X���[�h����ۑ�
        .SaveData = True              '�s�{�b�g�e�[�u���E���|�[�g�̒�`�ƃf�[�^��ۑ�
        .AdjustColumnWidth = False    '�񕝂���������
        .RefreshPeriod = 0            '�N�G���E�e�[�u���������X�V����Ԋu
        .PreserveColumnInfo = True    '�N�G���E�e�[�u�����X�V����邲�ƂɁA
        '��̕��בւ���t�B���^�[�ݒ褃��C�A�E�g�Ɋւ������ۑ�
    End With

End Sub

''
' �f�[�^�V�[�g�̏����ݒ�
'
' @param {Worksheet} targetSheet �Ώۂ̃V�[�g
'
Public Sub SetSheetStyle(ByRef targetSheet As Worksheet)

    With targetSheet
        Dim navyBlue As Long: navyBlue = RGB(0, 32, 96)
        .Tab.Color = navyBlue
        
        '�S�̂̐ݒ�
        With .Cells
            With .Font
                .Name = Me.Font
                .Size = Me.FontSize
            End With
            
            .RowHeight = Me.RowHeight
            .VerticalAlignment = Me.VerticalAlignment
            .NumberFormatLocal = Me.TypeOfNumbers
        End With
        
        '�w�b�_�[�̐ݒ�
        Dim sheetHeader As Range: Set sheetHeader = .Rows(1)
        sheetHeader.Font.Size = Me.HeaderFontSize
        Call SetColorFormula(sheetHeader)
        
        '�񕝂̐ݒ�
        With .Columns
            .ColumnWidth = Me.ColumnWidth_GridPaper
            .AutoFit
        End With
        
        '�񂲂Ƃ̐ݒ�
        Call SetColumnStyles(targetSheet)
        
        '�㏈��
        sheetHeader.HorizontalAlignment = xlLeft
        
        .Cells(2, 2).Select
        ActiveWindow.FreezePanes = True '���ɐݒ肳��Ă�����ς��Ȃ�
        
    End With

End Sub

''
' ��̐ݒ�
'
' @param {Worksheet} �Ώۂ̃V�[�g
'
Public Sub SetColumnStyles(ByRef targetSheet As Worksheet)

    '��̐ݒ�
    With targetSheet
        Dim i As Long
        For i = 1 To .Columns.Count
            With .Columns(i)
                If .Cells(1, 1).Value = "" Then Exit Sub
            
                '�񖼂����ė�̐ݒ���s��
                Dim isApplied As Boolean
                isApplied = applyColumnSettingsByHeader(.Columns(i))
                
                '���̑�
                If isApplied = False Then Call setColumnWidth(.Columns(i))
                
            End With
        Next i
    End With

End Sub

''
' ��̐ݒ�l��K�p
'
' @param {Range} targetColumn �K�p��
'
Private Function applyColumnSettingsByHeader(ByRef targetColumn As Range) As Boolean

    Dim headerString As String
    headerString = targetColumn.Cells(1, 1).Value

    Dim left2 As String:   left2 = Left(headerString, 2)
    Dim right1 As String: right1 = Right(headerString, 1)
    Dim right2 As String: right2 = Right(headerString, 2)
    Dim right3 As String: right3 = Right(headerString, 3)
    Dim right4 As String: right4 = Right(headerString, 4)

    With Me
        '�����iH��AM��j
        Select Case left2
            Case "H_", "��_"
                Call applyColumnSettings(targetColumn, xlRight, .TypeOfNumbers, .ColumnWidth_MinutesAndHour)
                applyColumnSettingsByHeader = True
                Exit Function
            
            Case "M_", "��_"
                Call applyColumnSettings(targetColumn, xlLeft, .TypeOfNumbers00, .ColumnWidth_MinutesAndHour)
                applyColumnSettingsByHeader = True
                Exit Function
        
        End Select
        
        Select Case True '�������^�[��
            '����
            Case right2 = "����"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfDateTime_WithDayOfTheWeek, .ColumnWidth_DateTime_WithDayOfTheWeek)
                applyColumnSettingsByHeader = True
                Exit Function

            Case right1 = "��"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfDate_WithDayOfTheWeek, .ColumnWidth_Date_WithDayOfTheWeek)
                applyColumnSettingsByHeader = True
                Exit Function
            
            'ID
            Case right2 = "ID"
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
        
            'No
            Case (right2 = "No"), (right3 = "No.")
                Call applyColumnSettings(targetColumn, xlRight, .TypeOfNumbers, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
        
            '�J�e�S��
            Case (right2 = "�J�e"), (right4 = "�J�e�S��")
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, canAutofit:=True)
                applyColumnSettingsByHeader = True
                Exit Function
            
            'True/False
            Case (left2 = "is"), (right1 = "�H"), (right1 = "?")
                Call applyColumnSettings(targetColumn, xlCenter, .TypeOfStandard, .ColumnWidth_Boolean)
                applyColumnSettingsByHeader = True
                Exit Function
                                        
        End Select

    End With

End Function

''
' ��̐ݒ��K�p����
'
' @param {Range}   targetColumn         �Y����
' @param {Long}    aHorizontalAlignment �l�̐����ʒu
' @param {String}  aNumberFormatLocal   �l�̏���
' @param {Long}    aColumnWidth         ��
' @param {Boolean} canAutofit           ���������������Ă悢��
'
Private Sub applyColumnSettings(ByRef targetColumn As Range, _
                                Optional ByVal aHorizontalAlignment As Long = xlRight, _
                                Optional ByVal aNumberFormatLocal As String, _
                                Optional ByVal aColumnWidth As Long, _
                                Optional ByVal canAutofit As Boolean = False)
    
    With targetColumn
        .HorizontalAlignment = aHorizontalAlignment
        
        If aNumberFormatLocal = "" Then aNumberFormatLocal = Me.TypeOfStandard
        .NumberFormatLocal = aNumberFormatLocal
        
        If canAutofit = False Then
            If IsEmpty(aColumnWidth) Then aColumnWidth = Me.ColumnWidth_NARROW
            .ColumnWidth = aColumnWidth
        
        Else
            .AutoFit
        
        End If
    
    End With

End Sub

''
' �������̕��ɒ�����
'
' @param {Range} targetColumn ����ς����
'
Private Sub setColumnWidth(ByRef targetColumn As Range)

    With targetColumn

        Select Case .ColumnWidth
            Case Is > Me.ColumnWidth_MAX
                .ColumnWidth = Me.ColumnWidth_MAX
            
            Case Is > Me.ColumnWidth_MIDDLE
                .ColumnWidth = Me.ColumnWidth_MIDDLE
            
            Case Is > Me.ColumnWidth_MIDDLE
                .ColumnWidth = Me.ColumnWidth_NARROW
            
        End Select
        
    End With

End Sub

''
' �����t�������̐ݒ�
' �i�֐��������̓w�b�_�[�ɐF�t������j
'
' @param {Range} targetRow �ݒ�̑Ώ۔͈�
'
Public Sub SetColorFormula(ByRef targetRow As Range)
    
    With targetRow
        .FormatConditions.Add Type:=xlExpression, _
                              Formula1:="=ISFORMULA(A$2)"
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With targetRow.FormatConditions(1)
            With .Font
                .Bold = True
                .Color = RGB(0, 0, 0) '��
            End With
            
            .Interior.Color = RGB(226, 239, 218) '�W����
            
            .StopIfTrue = False
        End With
    End With

End Sub