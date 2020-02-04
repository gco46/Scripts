Attribute VB_Name = "ShortcutsSubFunc"
Option Explicit

Public Sub GetHightOfMergedCells(a_rRange As Range, a_iArRow())
'
' AutoFitMergedCellsHeight() �̃T�u���[�`��
' �����Z���̍������擾
'
    Dim iRowCount
    Dim i
    
    ReDim a_iArRow(0)
    
    iRowCount = a_rRange.MergeArea.Rows.Count
    
    For i = 1 To iRowCount
        ReDim Preserve a_iArRow(i - 1)
        a_iArRow(i - 1) = a_rRange.MergeArea.Rows(i).RowHeight
    Next
End Sub

Public Function GetWidthOfMergedCells(a_rRange As Range)
'
' AutoFitMergedCellsHeight() �̃T�u���[�`��
' �����Z���̕����擾
'
    Dim iColCount
    Dim i
    Dim iWidth
    
    iColCount = a_rRange.MergeArea.Columns.Count
    iWidth = 0
    
    For i = 1 To iColCount
        iWidth = iWidth + a_rRange.MergeArea.Item(i).ColumnWidth
    Next
    
    GetWidthOfMergedCells = iWidth
End Function

Function GetAddressesOfMergedCells(a_rRange As Range, a_sArCell() As String) As Boolean
'
' AutoFitMergedCellsHeight() �̃T�u���[�`��
' �Z���͈͂̊e�Z�����W���擾
'
    Dim sAddress                        '// �Z���ʒu
    Dim v
    Dim sStartCell
    Dim sEndCell
    Dim iMergeCount
    Dim i
    
    '// �Z���͈͎擾
    iMergeCount = a_rRange.MergeArea.Count
    
    For i = 0 To iMergeCount - 1
        ReDim Preserve a_sArCell(i)
        
        '// Item��1�X�^�[�g
        a_sArCell(i) = a_rRange.MergeArea.Item(i + 1).Address(False, False, ReferenceStyle:=xlA1)
    Next
    
    GetAddressesOfMergedCells = True
End Function

Public Sub SetHightOfMergedCells(a_rRange As Range, a_iBeforeHeight, a_iArRow())
'
' AutoFitMergedCellsHeight() �̃T�u���[�`��
' �����Z���̍�����ݒ�
'
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
