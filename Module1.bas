Attribute VB_Name = "Module1"
Const SHEET = "��荞�݌���"
Const MAXCELL = 46
Public Type haifuData
    N As String
    N_name As String
End Type
Dim TUDUKI_LINE As Integer
'** 29  ������
Public Sub makeSeikyusyoOfNikkaHome()
    
    '�A�v���P�[�V�����`����ϊ�
    With Application
        .ReferenceStyle = xlR1C1
    End With
    '***�V�[�g�f�[�^���N���A
    Call clearData
    '***�f�[�^��ǂݍ���
    Call makeNikkaHomeSeikyuData
    '***0648M020�̎}�Ԃ��܂Ƃ߂�
    Call chgdata("�f�[�^")
    '***�f�[�^��ǂݍ���
    Call makeNikkaHomeSeikyuData_0648MA1X
    '*** �f�[�^�̕��בւ�
    Call sortData
    '***�V�[�g���폜����i�O��z�z�f�[�^���폜�j
    Call initDelSheets
    '***�z�z����擾
    'Call gethaifusaki(haifusaki)
    '***�z�z�p�V�[�g�𓾈Ӑ斈�ɍ쐬
    Call makeSheet
    With Application
        .ReferenceStyle = xlA1
    End With
    '***�e�V�[�g�̕s�v�������폜
    Call DelSpaceArea
    Call OutputSheetFile
    MsgBox "�I���܂���"
End Sub


'***�z�z�p�V�[�g���H��ʂɍ쐬
Private Sub makeSheet()
    Dim i As Integer
    Dim tokuiCd As String
    Dim intLine As Integer  '***�f�[�^�̃J�E���^�[
    Dim intSline As Integer '***�\���̍s�J�E���^�[
    Dim intUline As Integer '***����̍s�J�E���^�[
    Dim subTotal As Double    '***���ׂ̍��v�i�[
    Dim strKeyKojino As String  '***��r�L�[�̍H���m�n
    
    intLine = 2
    
    With Sheets("�f�[�^")
        
        Do While .Cells(intLine, 1) <> ""
            If Trim(tokuiCd) <> Trim(.Cells(intLine, 5)) Then
                tokuiCd = .Cells(intLine, 5)
                Sheets("�\��").Select
                Sheets("�\��").Copy After:=Sheets(4)
                ActiveSheet.Name = tokuiCd & "�\��"
                Sheets("����").Select
                Sheets("����").Copy After:=Sheets(4)
                ActiveSheet.Name = tokuiCd & "����"
                Call initset(tokuiCd & "�\��", intLine)
                intSline = 15
                intUline = 6
                tokuiCd = .Cells(intLine, 5)
            End If
            '**�\��
            Call setData(tokuiCd & "�\��", intLine, intSline)
            '***����
            Call setUchiwakeData(tokuiCd & "����", intLine, intUline, strKeyKojino, subTotal)
            intLine = intLine + 1
        Loop

    End With

End Sub

'****************************************************************************
'***�����l���Z�b�g����i�\���j
'****************************************************************************
Private Sub initset(ByVal sheetName As String, ByVal intLine As Integer)
    Dim intLineTitle As Integer
    
    Sheets(sheetName).Cells(8, 2) = Sheets("�f�[�^").Cells(intLine, 6)
    Sheets(sheetName).Cells(5, 18) = Format(DateAdd("M", -1, Date), "MM")
    
    '***�^�C�g��
    If InStr(Sheets("�f�[�^").Cells(intLine, 6), "��≮") > 0 Then
       Sheets(sheetName).Cells(6, 1) = "��≮�@������Ё@�䒆"
       Sheets(sheetName).Cells(8, 2) = Replace(Sheets("�f�[�^").Cells(intLine, 6), "��≮��", "")
    Else
        If Mid(Sheets("�f�[�^").Cells(intLine, 6), 1, 8) = "�j�b�J�z�[���֓�" Then
            Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������Ё@�䒆"
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets("�f�[�^").Cells(intLine, 6) & "�c�Ə�", "��", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "�֓�", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "�j�b�J�z�[��", "")
        Else
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets("�f�[�^").Cells(intLine, 6), "��", "")
            Sheets(sheetName).Cells(8, 2) = Replace(Sheets(sheetName).Cells(8, 2), "�j�b�J�z�[��", "")
        End If
    End If
    intLineTitle = 30
    Debug.Print Sheets("�f�[�^").Cells(intLine, 5)
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348M570" Then
        Debug.Print
    End If
    

    Do While Sheets("���j���[").Cells(intLineTitle, 2) <> ""
        If Sheets("���j���[").Cells(intLineTitle, 2) = Sheets("�f�[�^").Cells(intLine, 5) Then
            If Sheets("���j���[").Cells(intLineTitle, 1) <> "" Then
                Sheets(sheetName).Cells(6, 1) = Sheets("���j���[").Cells(intLineTitle, 1) & " �䒆"
            End If
            Exit Sub
        End If
        intLineTitle = intLineTitle + 1
    Loop


    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M870" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "������c�Ə�"
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M880" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�����}���i�c"
    End If
    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M890" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�������c�Ə�"
    End If
    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M900" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�������ǁi�c"
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME30" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮�@�쉡�l�X"
    End If
    
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮�@�������X"
    End If

    '***20200221
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME60" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��^���ƕ�"
    End If
    
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮�@�������X"
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "0348M710" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "���l�ۓy���J�c�Ə�"
    End If


    If Sheets("�f�[�^").Cells(intLine, 5) = "0348M690" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "�{��"
    End If

End Sub

'****************************************************************************
'***�f�[�^���Z�b�g����i�\���j
'****************************************************************************
Private Sub setData(ByVal sheetName As String, ByVal intLine As Integer, ByRef intDummy As Integer)
    Dim aTenmei As Variant
    Dim i As Integer
    Dim intSline As Integer
    Dim btrue As Boolean
    Dim atenmeiCHK As String

    intSline = 16
    btrue = False

    '***�ŏ���1�s�̋��z�����󔒂���Ȃ�������A�����H���ԍ���T���{���ꖼ����v����΂n�j
    '*cells(intSline,11) �͋��z��
    '*cells(intline,20) �͒���
    Do While Sheets(sheetName).Cells(intSline, 11) <> ""
        If Trim(Sheets(sheetName).Cells(intSline, 1)) = Trim(Sheets("�f�[�^").Cells(intLine, 20)) Then
            Sheets("�f�[�^").Cells(intLine, 29) = Replace(Sheets("�f�[�^").Cells(intLine, 29), "*", "/")    '**9/1�X�V
            If InStr(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^") > 0 Then
                    aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^")
                    atenmeiCHK = aTenmei(1)
            Else
                If InStr(Sheets("�f�[�^").Cells(intLine, 29), "/") > 0 Then
                    aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 29), "/")
                    atenmeiCHK = aTenmei(1)
                Else
                    atenmeiCHK = ""                                                                                 '**���ꖼ
                End If
            End If
            '*���ꖼ���������H�����Ȃ���z�����Z����B
            '*���ꖼ����H���ԍ��ɕύX����
            'If Trim(Sheets(sheetName).Cells(intSline, 3)) = Trim(atenmeiCHK) Then
            If Trim(Sheets(sheetName).Cells(intSline, 1)) = Trim(Sheets("�f�[�^").Cells(intLine, 20)) Then
                Sheets(sheetName).Cells(intSline, 11) = Sheets(sheetName).Cells(intSline, 11) + Sheets("�f�[�^").Cells(intLine, 19)
                btrue = True
            End If
        End If
        intSline = intSline + 1
        If intSline = 37 Then
            intSline = 42
        End If
        If intSline = 72 Then
            intSline = 77
        End If
        If intSline = 72 Then
            intSline = 77
        End If
        If intSline = 108 Then
            intSline = 111
        End If
    Loop
    
    If Not (btrue) Then
        Sheets(sheetName).Cells(intSline, 1) = Trim(Sheets("�f�[�^").Cells(intLine, 20))                        '**�H���ԍ��i���ԁj
        Sheets("�f�[�^").Cells(intLine, 29) = Replace(Sheets("�f�[�^").Cells(intLine, 29), "*", "/")    '**9/1�X�V
        If InStr(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^") > 0 Then
                aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^")
                Sheets(sheetName).Cells(intSline, 3) = aTenmei(1)                                                   '**���ꖼ
                Sheets(sheetName).Cells(intSline, 8) = aTenmei(0)                                                   '**�S��
        Else
            If InStr(Sheets("�f�[�^").Cells(intLine, 29), "/") > 0 Then
                aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 29), "/")
                Sheets(sheetName).Cells(intSline, 3) = aTenmei(1)                                                   '**���ꖼ
                Sheets(sheetName).Cells(intSline, 8) = aTenmei(0)                                                   '**�S��
            Else
                Sheets(sheetName).Cells(intSline, 8) = Sheets("�f�[�^").Cells(intLine, 29)                          '**���ꖼ
            End If
        End If
        Sheets(sheetName).Cells(intSline, 11) = Sheets("�f�[�^").Cells(intLine, 19)                             '**���z
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M870" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "������c�Ə�"
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M880" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�����}���i�c"
    End If
    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M890" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�������c�Ə�"
    End If
    If Sheets("�f�[�^").Cells(intLine, 5) = "1148M900" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���������@�䒆"
       Sheets(sheetName).Cells(8, 2) = "�������ǁi�c"
    End If
    
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME30" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮ �쉡�l�X"
    End If
    
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮ �������X�@"
    End If
    '***20200221
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME60" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��^���ƕ�"
    End If
    
    If Sheets("�f�[�^").Cells(intLine, 5) = "0348ME50" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "��≮�@�������X"
    End If

    If Sheets("�f�[�^").Cells(intLine, 5) = "0348M710" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "���l�ۓy���J�c�Ə�"
    End If


    If Sheets("�f�[�^").Cells(intLine, 5) = "0348M690" Then
       Sheets(sheetName).Cells(6, 1) = "�j�b�J�z�[���֓�������� �䒆"
       Sheets(sheetName).Cells(8, 2) = "�{��"
    End If
End Sub

'****************************************************************************
'***�f�[�^���Z�b�g����i����E���ׁj
'****************************************************************************
Private Sub setUchiwakeData(ByVal sheetName As String, _
                            ByVal intLine As Integer, _
                            ByRef intUline As Integer, _
                            ByRef strKeyKojino As String, _
                            ByRef subTotal As Double)
    
    Dim aTenmei As Variant
    '***�H����
    
    subTotal = subTotal + Sheets("�f�[�^").Cells(intLine, 18) * Sheets("�f�[�^").Cells(intLine, 17)
    Sheets(sheetName).Cells(intUline, 2) = Trim(Sheets("�f�[�^").Cells(intLine, 20))                              '**�H���ԍ��i����
    Sheets(sheetName).Cells(intUline, 1) = Mid(Sheets("�f�[�^").Cells(intLine, 21), 3, 2) & "/" _
                                         & Mid(Sheets("�f�[�^").Cells(intLine, 21), 5, 2)                       '***�o�ד�
    Sheets(sheetName).Cells(intUline + 1, 5) = Sheets("�f�[�^").Cells(intLine, 17)                              '***����
    Sheets(sheetName).Cells(intUline + 1, 6) = Sheets("�f�[�^").Cells(intLine, 18)                              '***�P��
    '***�W�v�������(�H���ԍ�,���t)
    If strKeyKojino <> Trim(Sheets("�f�[�^").Cells(intLine + 1, 20)) Or _
        Trim(Sheets("�f�[�^").Cells(intLine + 1, 5)) <> Trim(Sheets("�f�[�^").Cells(intLine, 5)) Or _
        Trim(Sheets("�f�[�^").Cells(intLine + 1, 21)) <> Trim(Sheets("�f�[�^").Cells(intLine, 21)) Then
        If subTotal <> 0 Then
            Sheets(sheetName).Cells(intUline + 1, 8) = subTotal
        End If
        subTotal = 0
        strKeyKojino = Trim(Sheets("�f�[�^").Cells(intLine + 1, 20))
    End If
    '***��E�𔻒�i�}�Ԃ�����Ȃ��j
    If Trim(Sheets("�f�[�^").Cells(intLine, 47)) = "" Then
        Sheets(sheetName).Cells(intUline, 5) = "��"
        Sheets(sheetName).Cells(intUline + 1, 2) = Sheets("�f�[�^").Cells(intLine, 15)
    Else
        Sheets(sheetName).Cells(intUline, 5) = "��"
        Sheets(sheetName).Cells(intUline + 1, 2) = Sheets("�f�[�^").Cells(intLine, 15)
    End If
    
    '***���ꖼ�ƒS����(2019.9.2 ���l�̓��e���󎚂���悤�ɕύX)
    If InStr(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^") > 0 Then
        aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 54) & Sheets("�f�[�^").Cells(intLine, 55), "�^")
        Sheets(sheetName).Cells(intUline, 3) = aTenmei(1)                                                       '**���ꖼ
        Sheets(sheetName).Cells(intUline, 4) = aTenmei(0)                                                       '**�S��
    Else
        If InStr(Sheets("�f�[�^").Cells(intLine, 29), "/") > 0 Then
            aTenmei = Split(Sheets("�f�[�^").Cells(intLine, 29), "/")
            Sheets(sheetName).Cells(intUline, 3) = aTenmei(1)                                                       '**���ꖼ
            Sheets(sheetName).Cells(intUline, 4) = aTenmei(0)                                                       '**�S��
        Else
            Sheets(sheetName).Cells(intUline, 4) = Sheets("�f�[�^").Cells(intLine, 29)                              '**���ꖼ
        End If
    End If
    intUline = intUline + 2
    If intUline = 56 Then
        intUline = 61
    End If

    If intUline = 111 Then
        intUline = 116
    End If

    If intUline = 166 Then
        intUline = 171
    End If
    
    If intUline = 221 Then
        intUline = 226
    End If
    If intUline = 276 Then
        intUline = 281
    End If
    If intUline = 331 Then
        intUline = 336
    End If
    If intUline = 386 Then
        intUline = intUline + 5
    End If
    If intUline = 441 Then
        intUline = intUline + 5
    End If
    If intUline = 496 Then
        intUline = intUline + 5
    End If
    If intUline = 551 Then
        intUline = intUline + 5
    End If
    'If intUline > 386 Then
        'If intUline Mod 55 = 0 Then
            'intUline = intUline + 5
        'End If
    'End If
        


End Sub
'****************************************************************************
'***�z�z��V�[�g�ɓ\��t����
'****************************************************************************
Private Sub PasteFactToSheet(ByVal sheetName As String, ByVal stN As String)
    Dim iRow As Integer
    Dim iColumn As Integer
    
    With Sheets(SHEET)
        .Cells(1, 1).AutoFilter Field:=3, Criteria1:=stN
        iColumn = .Cells(1, 1).End(xlToRight).Column
        iRow = .Cells(1, 1).End(xlDown).Row
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).Copy
        Sheets(sheetName).Paste
        '***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
        Call insFunc(sheetName)
    End With
End Sub
'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************
Private Sub insFunc(ByVal sheetName As String)
    Dim iniLine As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim i As Integer
    
    intLine = 1
    
    With Sheets(sheetName)
        iColumn = .Cells(1, 1).End(xlToRight).Column
        iRow = .Cells(1, 1).End(xlDown).Row
        '***��U�F�����Z�b�g����
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).Interior.ColorIndex = xlNone
    
        Do While .Cells(intLine, 1) <> ""
            If .Cells(intLine, 13) = "��s�m�F" Then
                '***�F������
                .Range(.Cells(intLine, 1), .Cells(intLine, MAXCELL)).Interior.ColorIndex = 35
                For i = 0 To 30
                    .Cells(intLine, 15 + i) = "=RC[-1]+R[-3]C-R[-2]C-R[-1]C"
                Next i

            End If
            intLine = intLine + 1
        Loop
        '***�����t�����ݒ������
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions.Delete
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="0"
        .Range(.Cells(1, 1), .Cells(iRow, iColumn)).FormatConditions(1).Interior.ColorIndex = 7
        .Cells.EntireColumn.AutoFit
    End With

End Sub
'****************************************************************************
'***�H��z�z�p�̃V�[�g���폜����
'****************************************************************************
Private Sub initDelSheets()
    Dim ws As Worksheet
    '***�폜�v���OFF�ɂ���
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> "�\��" And ws.Name <> "����" And ws.Name <> "�f�[�^" And ws.Name <> "���j���[" Then
            Sheets(ws.Name).Delete
        End If
    Next ws
    '***�폜�v���ON�ɂ���
    Application.DisplayAlerts = True

End Sub
'***�z�z����擾
Private Sub gethaifusaki(ByRef haifusaki() As haifuData)
    
    Dim i As Integer
    
    i = 0
    With Sheet2
        Do While .Cells(17 + i, 7) <> ""
           ReDim Preserve haifusaki(i)
            haifusaki(i).N = .Cells(17 + i, 7)
            haifusaki(i).N_name = .Cells(17 + i, 7 + 1)
            i = i + 1
        Loop
    End With
End Sub
'*************************************************************************************************
'�V�[�g���Ƀf�[�^�����ɍ��킹�ăy�[�W��ݒ肷��i����Ȃ��Ƃ���������j
'*************************************************************************************************
Public Sub DelSpaceArea()
    Dim intIdx As Integer       '�����p�C���f�b�N�X
    Dim intWksCnt As Integer    '�����p�J�E���^
    
    '�V�[�g���擾
    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
    '***OFF�ɂ���
    Application.DisplayAlerts = False
    
    '�V�[�g�����[�v
    For intIdx = 1 To intWksCnt
        '�ΏۃV�[�g���擾
        strWksnme = Worksheets(intIdx).Name
        If Right(strWksnme, 2) = "�\��" And strWksnme <> "�\��" Then
            If Sheets(strWksnme).Cells(42, 1) = "" Then
                Call Setpage1(strWksnme)
            Else
                If Sheets(strWksnme).Cells(77, 1) = "" Then
                    Call Setpage2(strWksnme)
                Else
                    If Sheets(strWksnme).Cells(111, 1) = "" Then
                        Call Setpage3(strWksnme)
                    Else
                        If Sheets(strWksnme).Cells(145, 1) = "" Then
                            Call Setpage4(strWksnme)
                        End If
                    End If

                End If
            End If

        End If
        If Right(strWksnme, 2) = "����" And strWksnme <> "����" Then
            '�P�y�[�W
            If Sheets(strWksnme).Cells(61, 1) = "" Then
                Call SetpageUchiwake1(strWksnme)
            Else
                '�Q�y�[�W
                If Sheets(strWksnme).Cells(116, 1) = "" Then
                    Call SetpageUchiwake2(strWksnme)
                Else
                    '�R�y�[�W
                    If Sheets(strWksnme).Cells(171, 1) = "" Then
                        Call SetpageUchiwake3(strWksnme)
                     Else
                        '4�y�[�W
                        If Sheets(strWksnme).Cells(226, 1) = "" Then
                            Call SetpageUchiwake4(strWksnme)
                        Else
                        '5�y�[�W
                            If Sheets(strWksnme).Cells(281, 1) = "" Then
                                Call SetpageUchiwake5(strWksnme)
                            Else
                            '6�y�[�W
                            If Sheets(strWksnme).Cells(336, 1) = "" Then
                                Call SetpageUchiwake6(strWksnme)
                            Else
                            '7�y�[�W
                            If Sheets(strWksnme).Cells(391, 1) = "" Then
                                Call SetpageUchiwake7(strWksnme)
                            Else
                            '8�y�[�W
                            If Sheets(strWksnme).Cells(446, 1) = "" Then
                                Call SetpageUchiwake8(strWksnme)
                            Else
                            '9�y�[�W
                            If Sheets(strWksnme).Cells(501, 1) = "" Then
                                Call SetpageUchiwake9(strWksnme)
                            Else
                                MsgBox ("�y�[�W�z��O" & strWksnme)
                            End If
                            End If
                            End If
                            End If
                            End If
                            

                        End If
                    End If
                End If
            End If
        End If
    Next
    '***OFF�ɂ���
    Application.DisplayAlerts = True

End Sub
''*************************************************************************************************
''�V�[�g���ʂɃt�@�C���Ƃ��ďo�͂���
''*************************************************************************************************
'Public Sub OutputSheetFile()
'    Dim intIdx As Integer       '�����p�C���f�b�N�X
'    Dim intWksCnt As Integer    '�����p�J�E���^
'    Dim objWks As Object        '�V�[�g�쐬�p�I�u�W�F�N�g
'    Dim strWbkNme As String     'Excel���[�N�u�b�N��(�g���q�܂܂�)
'    Dim strWbkDir As String     'Excel���[�N�u�b�N�ۑ��ꏊ
'    Dim strWksnme As String     '�V�[�g��
'
'    'Excel���[�N�u�b�N�̏��擾
'    strWbkDir = Application.ActiveWorkbook.Path
'    strWbkNme = Application.ActiveWorkbook.Name
'    If Right(strWbkNme, Len(".xls")) = ".xls" Then
'        strWbkNme = Left(strWbkNme, Len(strWbkNme) - Len(".xls"))
'    End If
'
'    '�V�[�g���擾
'    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
'    '***OFF�ɂ���
'    Application.DisplayAlerts = False
'
'    '�V�[�g�����[�v
'    For intIdx = 1 To intWksCnt
'        '�ΏۃV�[�g���擾
'        strWksnme = Worksheets(intIdx).Name
'        If Right(strWksnme, 2) = "�\��" And strWksnme <> "�\��" Then
'            '�V�[�g�̃R�s�[
'            Worksheets(intIdx).Copy
'            '�t�@�C���ۑ�
'            ActiveWorkbook.SaveAs Filename:= _
'                strWbkDir & "\" & strWksnme & ".xls", _
'                FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
'                ReadOnlyRecommended:=False, CreateBackup:=False
'                Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "����").Copy Before:=Sheets(1)
'            ActiveWorkbook.Save
'
'            'Sheets(Array(Left(strWksnme, 8) & "�\��", Left(strWksnme, 8) & "����")).Select
'            Sheets(Array(Left(strWksnme, 8) & "����", Left(strWksnme, 8) & "�\��")).Select
'            'Sheets(Left(strWksNme, 8) & "�\��").Activate
'            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'                "C:\Users\KATOTO\Documents\���[�����M\" & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
'                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
'            ActiveWindow.Close
'        End If
'    Next
'    '***OFF�ɂ���
'    Application.DisplayAlerts = True
'
'End Sub
'*************************************************************************************************
'�V�[�g���ʂɃt�@�C���Ƃ��ďo�͂���
'*************************************************************************************************
Public Sub OutputSheetFile()
    Dim intIdx As Integer       '�����p�C���f�b�N�X
    Dim intWksCnt As Integer    '�����p�J�E���^
    Dim objWks As Object        '�V�[�g�쐬�p�I�u�W�F�N�g
    Dim strWbkNme As String     'Excel���[�N�u�b�N��(�g���q�܂܂�)
    Dim strWbkDir As String     'Excel���[�N�u�b�N�ۑ��ꏊ
    Dim strWksnme As String     '�V�[�g��
    Dim outFolder As String
    Dim outFolder2 As String

    'Excel���[�N�u�b�N�̏��擾
    strWbkDir = Application.ActiveWorkbook.Path
    strWbkNme = Application.ActiveWorkbook.Name
    If Right(strWbkNme, Len(".xls")) = ".xls" Then
        strWbkNme = Left(strWbkNme, Len(strWbkNme) - Len(".xls"))
    End If

    'outFolder = "\\hob1sv07ap\���ϋ��L$\��c�p\�j�b�J�z�[��������\��≮�֓�\"
    'outFolder1 = "\\hob1sv07ap\���ϋ��L$\��c�p\�j�b�J�z�[��������\�j�b�J�z�[���֓�\"
    'outFolder2 = "\\hob1sv07ap\���ϋ��L$\��c�p\�j�b�J�z�[��������\���̑�\"
    'outFolder = "C:\Users\KATOTO\Documents\�j�b�J�z�[��\�e�X�g\"
    'outFolder2 = "C:\Users\KATOTO\Documents\�j�b�J�z�[��\�e�X�g\"
    outFolder = "\\HOB1SV03FS\�̔��̑S�����L\�����Ǘ�����\���j�b�J�z�[��������\��≮�֓�\"
    outFolder1 = "\\HOB1SV03FS\�̔��̑S�����L\�����Ǘ�����\���j�b�J�z�[��������\�j�b�J�z�[���֓�\"
    outFolder2 = "\\HOB1SV03FS\�̔��̑S�����L\�����Ǘ�����\���j�b�J�z�[��������\���̑�\"
    'outFolder = "C:\Users\KATOTO\Documents\�j�b�J�z�[��\�e�X�g\��≮�֓�\"
    'outFolder1 = "C:\Users\KATOTO\Documents\�j�b�J�z�[��\�e�X�g\�j�b�J�z�[���֓�\"
    'outFolder2 = "C:\Users\KATOTO\Documents\�j�b�J�z�[��\�e�X�g\���̑�\"
    '�V�[�g���擾
    intWksCnt = Excel.ActiveWorkbook.Worksheets.Count
    '***OFF�ɂ���
    Application.DisplayAlerts = False

    '�V�[�g�����[�v
    For intIdx = 1 To intWksCnt
        '�ΏۃV�[�g���擾
        strWksnme = Worksheets(intIdx).Name
        If Right(strWksnme, 2) = "����" And strWksnme <> "����" Then
            '�V�[�g�̃R�s�[
            Worksheets(intIdx).Copy
            If Left(Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Cells(6, 1), 3) = "��≮" And Left(strWksnme, 2) = "03" Then
                '�t�@�C���ۑ�
                ActiveWorkbook.SaveAs Filename:= _
                    outFolder & strWksnme & ".xls", _
                    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                    ReadOnlyRecommended:=False, CreateBackup:=False
                    Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Copy Before:=Sheets(1)
                ActiveWorkbook.Save
    
                Sheets(Array(Left(strWksnme, 8) & "�\��", Left(strWksnme, 8) & "����")).Select
                'Sheets(Array(Left(strWksnme, 8) & "����", Left(strWksnme, 8) & "�\��")).Select
                'Sheets(Left(strWksNme, 8) & "�\��").Activate
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    outFolder & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                ActiveWindow.Close
            Else
                If (Left(Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Cells(6, 1), 6) = "�j�b�J�z�[��" Or _
                Left(Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Cells(6, 1), 11) = "������� �j�b�J�z�[��") And Left(strWksnme, 2) = "03" Then
                    '�t�@�C���ۑ�
                    ActiveWorkbook.SaveAs Filename:= _
                        outFolder1 & strWksnme & ".xls", _
                        FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                        ReadOnlyRecommended:=False, CreateBackup:=False
                        Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Copy Before:=Sheets(1)
                    ActiveWorkbook.Save
        
                    Sheets(Array(Left(strWksnme, 8) & "�\��", Left(strWksnme, 8) & "����")).Select
                    'Sheets(Array(Left(strWksnme, 8) & "����", Left(strWksnme, 8) & "�\��")).Select
                    'Sheets(Left(strWksNme, 8) & "�\��").Activate
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                        outFolder1 & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                    ActiveWindow.Close
                Else
                    '�t�@�C���ۑ�
                    ActiveWorkbook.SaveAs Filename:= _
                        outFolder2 & strWksnme & ".xls", _
                        FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                        ReadOnlyRecommended:=False, CreateBackup:=False
                        Workbooks("�j�b�J�z�[���������쐬.xlsm").Sheets(Left(strWksnme, 8) & "�\��").Copy Before:=Sheets(1)
                    ActiveWorkbook.Save
        
                    Sheets(Array(Left(strWksnme, 8) & "�\��", Left(strWksnme, 8) & "����")).Select
                    'Sheets(Array(Left(strWksnme, 8) & "����", Left(strWksnme, 8) & "�\��")).Select
                    'Sheets(Left(strWksNme, 8) & "�\��").Activate
                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                        outFolder2 & Left(strWksnme, 8) & ".pdf", Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=Ture, OpenAfterPublish:=False
                    ActiveWindow.Close
                End If
            End If
        End If
    Next
    '***OFF�ɂ���
    Application.DisplayAlerts = True

End Sub
'*************************************************************************************************
'�����f�[�^��ǂݍ���
'*************************************************************************************************
Public Sub makeNikkaHomeSeikyuData()
        
    Dim csvData As Variant
    Dim FSO As Object
    Dim intLine As Integer
    Dim intRow As Integer
    Dim i As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    intLine = 2
    
    Set csvfile = FSO.Opentextfile("C:\Users\KATOTO\Documents\�j�b�J�z�[��\NOHINPCX.CSV")
    'Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
    With csvfile
        Do Until .atendofstream
            csvData = Split(Replace(.readline, """", ""), ",")
            intRow = 1
            For i = 0 To UBound(csvData)
                Sheets("�f�[�^").Cells(intLine, intRow) = csvData(i)
                intRow = intRow + 1
            Next i
            intLine = intLine + 1
        Loop
        .Close
    End With
    
    If Sheets("���j���[").Cells(1, 1) = "�l��" Then
        
        '***�l���f�[�^
        Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
        With csvfile
            Do Until .atendofstream
                csvData = Split(Replace(.readline, """", ""), ",")
                intRow = 1
                For i = 0 To UBound(csvData)
                    Sheets("�f�[�^").Cells(intLine, intRow) = csvData(i)
                    intRow = intRow + 1
                Next i
                intLine = intLine + 1
            Loop
            .Close
        End With
    End If
    TUDUKI_LINE = intLine
End Sub
'*************************************************************************************************
'�����f�[�^��ǂݍ���
'2017.12.1 ���É��̗v�]��0648MA11,12,13��0648MA10�ɂ܂Ƃ߂Ă��A�e���ׂ͏o��
'*************************************************************************************************
Public Sub makeNikkaHomeSeikyuData_0648MA1X()
        
    Dim csvData As Variant
    Dim FSO As Object
    Dim intLine As Integer
    Dim intRow As Integer
    Dim i As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    intLine = TUDUKI_LINE
    
    Set csvfile = FSO.Opentextfile("C:\Users\KATOTO\Documents\�j�b�J�z�[��\NOHINPCX.CSV")
    With csvfile
        Do Until .atendofstream
            csvData = Split(Replace(.readline, """", ""), ",")
            intRow = 1
            If csvData(4) = "0648MA11" Or csvData(4) = "0648MA12" Or csvData(4) = "0648MA14" Or _
               csvData(4) = "0848ME21" Or csvData(4) = "0848ME22" Or csvData(4) = "0848ME23" Or csvData(4) = "0848ME24" Then
                For i = 0 To UBound(csvData)
                    Sheets("�f�[�^").Cells(intLine, intRow) = csvData(i)
                    intRow = intRow + 1
                Next i
                intLine = intLine + 1
            End If
        Loop
        .Close
    End With
    
    If Sheets("���j���[").Cells(1, 1) = "�l��" Then
        
        '***�l���f�[�^
        Set csvfile = FSO.Opentextfile("E:\NOHINPCX.CSV")
        With csvfile
            Do Until .atendofstream
                csvData = Split(Replace(.readline, """", ""), ",")
                intRow = 1
                If csvData(4) = "0648MA11" Or csvData(4) = "0648MA12" Or csvData(4) = "0648MA14" Then
                    For i = 0 To UBound(csvData)
                        Sheets("�f�[�^").Cells(intLine, intRow) = csvData(i)
                        intRow = intRow + 1
                    Next i
                    intLine = intLine + 1
                End If
            Loop
            .Close
        End With
    End If
    TUDUKI_LINE = intLine
End Sub

'*************************************************************************************************
'�ǂݍ��񂾃f�[�^��
'*************************************************************************************************
Sub clearData()
'
    Sheets("�f�[�^").Select
    With Sheets("�f�[�^")
        Rows("2:2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp
    End With
    
End Sub
Sub CopyHyosi(ByVal CopySheetname As String, ByVal PasteSheetname As String)

    Sheets(CopySheetname).Select
    Cells.Select
    Cells.Copy
    Sheets(PasteSheetname).Select
    Sheets(PasteSheetname).Cells.Paste
End Sub
Sub Macro3()
'
' Macro3 Macro
'

'
    Cells.Select
End Sub

Sub Macro5()
'
' Macro5 Macro
'

'
    Sheets("�\��").Select
    Sheets("�\��").Copy After:=Sheets(5)
End Sub
Sub Macro6()
'
' Macro6 Macro
'

'
    Sheets("1148M240����").Select
    Sheets("1148M240����").Copy Before:=Workbooks("20160229QWE�쐬.xlsx").Sheets(1)
End Sub
Sub sortData()
'
' Macro1 Macro
'

'�@�@���Ӑ�A���t�A���ԍ�
    ActiveWorkbook.Worksheets("�f�[�^").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�f�[�^").Sort.SortFields.Add Key:=Range("E2:E2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�f�[�^").Sort.SortFields.Add Key:=Range("U2:U2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�f�[�^").Sort.SortFields.Add Key:=Range("T2:T2500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�f�[�^").Sort.SortFields.Add Key:=Range("Q2:Q2500"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�f�[�^").Sort
        .SetRange Range("A1:BC2500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************
Private Sub chgdata(ByVal sheetName As String)
    Dim iniLine As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim i As Integer
    
    intLine = 1
    
    With Sheets(sheetName)
    
        Do While .Cells(intLine, 5) <> ""
            If Left(.Cells(intLine, 5), 7) = "0648M02" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0648M020"
            End If
            '***20171201 �܂Ƃ߈˗�
            If Left(.Cells(intLine, 5), 8) = "0648MA11" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA15" Or _
               Left(.Cells(intLine, 5), 8) = "0648M011" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA12" Or _
               Left(.Cells(intLine, 5), 8) = "0648MA14" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0648MA10"
            End If
            '***20181101 �܂Ƃ߈˗�
            If Left(.Cells(intLine, 5), 8) = "0848ME20" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME21" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME22" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME23" Or _
               Left(.Cells(intLine, 5), 8) = "0848ME24" Then
                Debug.Print .Cells(intLine, 5) & " "; intLine
                .Cells(intLine, 5) = "0848ME20"
                .Cells(intLine, 6) = "�G�k�X�e�[�W(��)�����{"
            End If
            intLine = intLine + 1
        Loop
    End With

End Sub

'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************

Sub Setpage1(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$37"
End Sub

'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************

Sub Setpage2(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$72"
End Sub

'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************

Sub Setpage3(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$107"
End Sub

'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************

Sub Setpage4(ByVal strWksnme As String)
'
' Macro4 Macro
'

'
     Sheets(strWksnme).Activate
'    Range("A37:U38").Select
'    Selection.Copy
'    Range("A37:U38").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'    Range("A10:H13").Select
'    Selection.Copy
'    Range("A10:H13").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'    Rows("39:90000").Select
'    Range(Selection, Selection.End(xlDown)).Select
'
'    Selection.Delete Shift:=xlUp
    
    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$U$141"
End Sub


'****************************************************************************
'***�z�z��V�[�g�ɓ\��t�������ƂɁA�֐������Ă���
'****************************************************************************

Sub SetpageUchiwake1(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$55"
End Sub

Sub SetpageUchiwake2(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$110"
End Sub

Sub SetpageUchiwake3(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$165"
End Sub


Sub SetpageUchiwake4(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$220"
End Sub


Sub SetpageUchiwake5(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$275"
End Sub

Sub SetpageUchiwake6(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$330"
End Sub

Sub SetpageUchiwake7(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$385"
End Sub

Sub SetpageUchiwake8(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$440"
End Sub

Sub SetpageUchiwake9(ByVal strWksnme As String)
     Sheets(strWksnme).Activate

    Cells(1, 1).Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$H$495"
End Sub
