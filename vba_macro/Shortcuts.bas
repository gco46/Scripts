Attribute VB_Name = "Shortcuts"
Option Explicit

Sub ShowMultiBook()
Attribute ShowMultiBook.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' �J���Ă���u�b�N�����ɕ��ׂĕ\�� (ctrl + shift + a)
'
    Windows.Arrange ArrangeStyle:=xlArrangeStyleVertical
End Sub

Sub ShowOneBook()
Attribute ShowOneBook.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' �A�N�e�B�u�ȃu�b�N��S��ʕ\�� (ctrl + shift + s)
'
    ActiveWindow.WindowState = xlMaximized
End Sub

Sub ReDrawBorders()
Attribute ReDrawBorders.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �I��͈͂̃Z�����͂��悤�Ɍr�������� (ctrl + q)
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub


Sub NoPaint()
Attribute NoPaint.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' �I���Z����"�h��Ԃ��Ȃ�"�ɂ��� (ctrl + shift + n)
'

    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub YellowPaint()
Attribute YellowPaint.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' �I���Z�������F�ɓh��Ԃ� (ctrl + shift + y)
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub AutoFitMergedCellsHeight()
Attribute AutoFitMergedCellsHeight.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' �����Z���̍������������� (ctrl + shift + r)
' Margin �ŃZ���̍����𒲐��\
'
    Dim Margin                  As Integer      ' �����������}�[�W��[px]
    Margin = 3

    Dim r                       As Range
    Dim iBondWidth                              '// �������̃Z����
    Dim iStartCellWidth                         '// �����ΏۃZ���̕�
    Dim sBeforeAddress                          '// �������̌����ΏۃZ��
    Dim iBeforeHeight                           '// �ŏI�I�ɐݒ肷�鍂��
    Dim sArCell()               As String       '// �Z���͈�
    Dim bRet                    As Boolean      '// �߂�l
    Dim i                                       '// ���[�v�J�E���^
    Dim iArCount                                '// �z��v�f��
    Dim sNowAddress                             '// ���݃Z�����W
    Dim bExistFlg               As Boolean      '// �z����ɃZ�������݂��Ă��邩����t���O�iTrue�F���݂���AFalse�F���݂��Ȃ��j
    Dim iArRow()                                '// �����s�̌������̊e�s�̍���

    ReDim sArCell(0)
    
    Application.ScreenUpdating = False
    
    For Each r In Selection
        '// �������̕����擾
        iBondWidth = GetWidthOfMergedCells(r)
        
        '// �������̊e�Z���̍������擾
        Call GetHightOfMergedCells(r, iArRow)
        
        iArCount = UBound(sArCell)
        
        '// �Z���z����Ɍ����[�v�̃Z��������Ώ����ΏہB�����łȂ���Ύ���Selection�����Ȃ̂Ō㑱�����Ŕz���蒼���B
        bExistFlg = False
        sNowAddress = r.Address(False, False)
        For i = 0 To iArCount
            If (sNowAddress = sArCell(i)) Then
                bExistFlg = True
                Exit For
            End If
        Next
        
        '// �z����Ɍ����[�v�̃Z�����Ȃ��ꍇ�i�����Z�������[�v�ŕς�����ꍇ�j
        If (bExistFlg = False) Then
            '// �Z���͈͂̑S�Z����z��Ŏ擾
            bRet = GetAddressesOfMergedCells(r, sArCell)
        End If
        
        '// �����Z���̏ꍇ
        If (sNowAddress = sArCell(0)) Then
            iStartCellWidth = r.ColumnWidth
            
            '// �������̌����ΏۃZ�����擾
            sBeforeAddress = r.MergeArea.Address(False, False, ReferenceStyle:=xlA1)
            
            '// ����������
            r.UnMerge
            
            '// �������̃Z�����܂Ŋg������
            r.ColumnWidth = iBondWidth
            
            '// �܂�Ԃ�ON
            r.WrapText = True
            
            '// �K�v�ȍ������擾
            r.EntireRow.AutoFit
            
            '// �Z���������擾
            iBeforeHeight = r.RowHeight + Margin
            
            '// �Č���
            Range(sBeforeAddress).Merge
            
            '// ������̃Z�������̃Z�����ɖ߂�
            r.MergeArea.Item(1).ColumnWidth = iStartCellWidth
            
            '// ������̃Z���̍�����ݒ�
            Call SetHightOfMergedCells(r, iBeforeHeight, iArRow)
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub


Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' �I���Z���̉���ɍ��킹�ăI�[�g�t�B�� (ctrl + shift + r)
'

    Dim myClm As Integer
    Dim myRow As Long
    Dim myRange1 As String
    Dim myRange2 As String

    myClm = Selection.Column
    myRow = ActiveSheet.Cells(Rows.Count, myClm).End(xlUp).Row
    If myClm <> 1 Then
        myRow = ActiveSheet.Cells(Rows.Count, myClm - 1).End(xlUp).Row
    Else
        myRow = ActiveSheet.Cells(Rows.Count, myClm + 1).End(xlUp).Row
    End If
    myRange1 = Selection.Address
    myRange2 = ActiveSheet.Cells(myRow, myClm).Address
    myRange2 = myRange1 & ":" & myRange2
    ActiveSheet.Range(myRange1).AutoFill Destination:=ActiveSheet.Range(myRange2)

End Sub
