Attribute VB_Name = "ShortcutMacro"
Option Explicit

Sub ShowMultiBook()
Attribute ShowMultiBook.VB_ProcData.VB_Invoke_Func = "A\n14"
    Windows.Arrange ArrangeStyle:=xlArrangeStyleVertical
End Sub

Sub ShowOneBook()
Attribute ShowOneBook.VB_ProcData.VB_Invoke_Func = "S\n14"
    ActiveWindow.WindowState = xlMaximized
End Sub

Sub ReDrawBorders()
Attribute ReDrawBorders.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' ReDrawBorders Macro
'

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
' NoPaint Macro
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
' Paint selected cells yellow
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Sub AutoFitConcatCellHeight()
Attribute AutoFitConcatCellHeight.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' �����Z���̍�������������
' Margin �ŃZ���̍����𒲐��\
'
    Dim Margin                  As Integer      ' �����������}�[�W��
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
        iBondWidth = �����Z���̕�(r)
        
        '// �������̊e�Z���̍������擾
        Call �����Z���̍������擾(r, iArRow)
        
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
            bRet = �Z���͈͂̊e�Z�����W���擾(r, sArCell)
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
            Call �����Z���̍�����ݒ�(r, iBeforeHeight, iArRow)
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub

Public Sub �����Z���̍������擾(a_rRange As Range, a_iArRow())
    Dim iRowCount
    Dim i
    
    ReDim a_iArRow(0)
    
    iRowCount = a_rRange.MergeArea.Rows.Count
    
    For i = 1 To iRowCount
        ReDim Preserve a_iArRow(i - 1)
        a_iArRow(i - 1) = a_rRange.MergeArea.Rows(i).RowHeight
    Next
End Sub

Public Function �����Z���̕�(a_rRange As Range)
    Dim iColCount
    Dim i
    Dim iWidth
    
    iColCount = a_rRange.MergeArea.Columns.Count
    iWidth = 0
    
    For i = 1 To iColCount
        iWidth = iWidth + a_rRange.MergeArea.Item(i).ColumnWidth
    Next
    
    �����Z���̕� = iWidth
End Function

Function �Z���͈͂̊e�Z�����W���擾(a_rRange As Range, a_sArCell() As String) As Boolean
    Dim sAddress                        '// �Z���ʒu
    Dim v
    Dim sStartCell
    Dim sEndCell
    Dim iMergeCount
    Dim i
    
    '// �Z���͈͎擾
    sAddress = a_rRange.MergeArea.Address(False, False, ReferenceStyle:=xlA1)
    iMergeCount = a_rRange.MergeArea.Count
    
    For i = 0 To iMergeCount - 1
        ReDim Preserve a_sArCell(i)
        
        '// Item��1�X�^�[�g
        a_sArCell(i) = a_rRange.MergeArea.Item(i + 1).Address(False, False, ReferenceStyle:=xlA1)
    Next
    
    �Z���͈͂̊e�Z�����W���擾 = True
End Function

Public Sub �����Z���̍�����ݒ�(a_rRange As Range, a_iBeforeHeight, a_iArRow())
    Dim iSumRow
    Dim i
    Dim iCount
    Dim iFirstRow
    Dim iRemainRow
    
    iSumRow = 0
    iCount = UBound(a_iArRow)
    
    '// �P��s�̏ꍇ
    If (iCount = 0) Then
        a_rRange.RowHeight = a_iBeforeHeight
        Exit Sub
    End If
    
    '// �ȉ������s�̏ꍇ
    
    For i = 0 To iCount
        iSumRow = iSumRow + a_iArRow(i)
    Next
    
    '// �擪�s�̍���
    iFirstRow = a_iArRow(0)
    '// �c��̍s�̍���
    iRemainRow = iSumRow - iFirstRow
    
    '// �������̍��������̍�����荂���ꍇ
    If (a_iBeforeHeight > iSumRow) Then
        a_rRange.RowHeight = a_iBeforeHeight - iRemainRow
    Else
        a_rRange.RowHeight = iFirstRow
    End If
End Sub


Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' �I���Z���̉���ɍ��킹�ăI�[�g�t�B��
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
