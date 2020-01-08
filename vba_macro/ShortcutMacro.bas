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
' 結合セルの高さを自動調節
' Margin でセルの高さを調整可能
'
    Dim Margin                  As Integer      ' 自動調整時マージン
    Margin = 3

    Dim r                       As Range
    Dim iBondWidth                              '// 結合時のセル幅
    Dim iStartCellWidth                         '// 処理対象セルの幅
    Dim sBeforeAddress                          '// 結合時の結合対象セル
    Dim iBeforeHeight                           '// 最終的に設定する高さ
    Dim sArCell()               As String       '// セル範囲
    Dim bRet                    As Boolean      '// 戻り値
    Dim i                                       '// ループカウンタ
    Dim iArCount                                '// 配列要素数
    Dim sNowAddress                             '// 現在セル座標
    Dim bExistFlg               As Boolean      '// 配列内にセルが存在しているか判定フラグ（True：存在する、False：存在しない）
    Dim iArRow()                                '// 複数行の結合時の各行の高さ

    ReDim sArCell(0)
    
    Application.ScreenUpdating = False
    
    For Each r In Selection
        '// 結合時の幅を取得
        iBondWidth = 結合セルの幅(r)
        
        '// 結合時の各セルの高さを取得
        Call 結合セルの高さを取得(r, iArRow)
        
        iArCount = UBound(sArCell)
        
        '// セル配列内に現ループのセルがあれば処理対象。そうでなければ次のSelection処理なので後続処理で配列取り直し。
        bExistFlg = False
        sNowAddress = r.Address(False, False)
        For i = 0 To iArCount
            If (sNowAddress = sArCell(i)) Then
                bExistFlg = True
                Exit For
            End If
        Next
        
        '// 配列内に現ループのセルがない場合（結合セルがループで変わった場合）
        If (bExistFlg = False) Then
            '// セル範囲の全セルを配列で取得
            bRet = セル範囲の各セル座標を取得(r, sArCell)
        End If
        
        '// 結合セルの場合
        If (sNowAddress = sArCell(0)) Then
            iStartCellWidth = r.ColumnWidth
            
            '// 結合時の結合対象セルを取得
            sBeforeAddress = r.MergeArea.Address(False, False, ReferenceStyle:=xlA1)
            
            '// 結合を解除
            r.UnMerge
            
            '// 結合時のセル幅まで拡張する
            r.ColumnWidth = iBondWidth
            
            '// 折り返しON
            r.WrapText = True
            
            '// 必要な高さを取得
            r.EntireRow.AutoFit
            
            '// セル高さを取得
            iBeforeHeight = r.RowHeight + Margin
            
            '// 再結合
            Range(sBeforeAddress).Merge
            
            '// 結合後のセルを元のセル幅に戻す
            r.MergeArea.Item(1).ColumnWidth = iStartCellWidth
            
            '// 結合後のセルの高さを設定
            Call 結合セルの高さを設定(r, iBeforeHeight, iArRow)
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub

Public Sub 結合セルの高さを取得(a_rRange As Range, a_iArRow())
    Dim iRowCount
    Dim i
    
    ReDim a_iArRow(0)
    
    iRowCount = a_rRange.MergeArea.Rows.Count
    
    For i = 1 To iRowCount
        ReDim Preserve a_iArRow(i - 1)
        a_iArRow(i - 1) = a_rRange.MergeArea.Rows(i).RowHeight
    Next
End Sub

Public Function 結合セルの幅(a_rRange As Range)
    Dim iColCount
    Dim i
    Dim iWidth
    
    iColCount = a_rRange.MergeArea.Columns.Count
    iWidth = 0
    
    For i = 1 To iColCount
        iWidth = iWidth + a_rRange.MergeArea.Item(i).ColumnWidth
    Next
    
    結合セルの幅 = iWidth
End Function

Function セル範囲の各セル座標を取得(a_rRange As Range, a_sArCell() As String) As Boolean
    Dim sAddress                        '// セル位置
    Dim v
    Dim sStartCell
    Dim sEndCell
    Dim iMergeCount
    Dim i
    
    '// セル範囲取得
    sAddress = a_rRange.MergeArea.Address(False, False, ReferenceStyle:=xlA1)
    iMergeCount = a_rRange.MergeArea.Count
    
    For i = 0 To iMergeCount - 1
        ReDim Preserve a_sArCell(i)
        
        '// Itemは1スタート
        a_sArCell(i) = a_rRange.MergeArea.Item(i + 1).Address(False, False, ReferenceStyle:=xlA1)
    Next
    
    セル範囲の各セル座標を取得 = True
End Function

Public Sub 結合セルの高さを設定(a_rRange As Range, a_iBeforeHeight, a_iArRow())
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


Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' 選択セルの横列に合わせてオートフィル
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
