Attribute VB_Name = "ShortcutsSubFunc"
Option Explicit

Public Sub GetHightOfMergedCells(a_rRange As Range, a_iArRow())
'
' AutoFitMergedCellsHeight() のサブルーチン
' 結合セルの高さを取得
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
' AutoFitMergedCellsHeight() のサブルーチン
' 結合セルの幅を取得
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
' AutoFitMergedCellsHeight() のサブルーチン
' セル範囲の各セル座標を取得
'
    Dim sAddress                        '// セル位置
    Dim v
    Dim sStartCell
    Dim sEndCell
    Dim iMergeCount
    Dim i
    
    '// セル範囲取得
    iMergeCount = a_rRange.MergeArea.Count
    
    For i = 0 To iMergeCount - 1
        ReDim Preserve a_sArCell(i)
        
        '// Itemは1スタート
        a_sArCell(i) = a_rRange.MergeArea.Item(i + 1).Address(False, False, ReferenceStyle:=xlA1)
    Next
    
    GetAddressesOfMergedCells = True
End Function

Public Sub SetHightOfMergedCells(a_rRange As Range, a_iBeforeHeight, a_iArRow())
'
' AutoFitMergedCellsHeight() のサブルーチン
' 結合セルの高さを設定
'
    Dim iSumRow
    Dim i
    Dim iCount
    Dim iFirstRow
    Dim iRemainRow
    
    iSumRow = 0
    iCount = UBound(a_iArRow)
    
    '// 単一行の場合
    If (iCount = 0) Then
        a_rRange.RowHeight = a_iBeforeHeight
        Exit Sub
    End If
    
    '// 以下複数行の場合
    
    For i = 0 To iCount
        iSumRow = iSumRow + a_iArRow(i)
    Next
    
    '// 先頭行の高さ
    iFirstRow = a_iArRow(0)
    '// 残りの行の高さ
    iRemainRow = iSumRow - iFirstRow
    
    '// 結合時の高さが元の高さより高い場合
    If (a_iBeforeHeight > iSumRow) Then
        a_rRange.RowHeight = a_iBeforeHeight - iRemainRow
    Else
        a_rRange.RowHeight = iFirstRow
    End If
End Sub
