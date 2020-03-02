Attribute VB_Name = "Shortcuts"
Option Explicit

Sub ShowMultiBook()
Attribute ShowMultiBook.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' 開いているブックを横に並べて表示 (ctrl + shift + a)
'
    Windows.Arrange ArrangeStyle:=xlArrangeStyleVertical
End Sub

Sub ShowOneBook()
Attribute ShowOneBook.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' アクティブなブックを全画面表示 (ctrl + shift + s)
'
    ActiveWindow.WindowState = xlMaximized
End Sub

Sub ReDrawBorders()
Attribute ReDrawBorders.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 選択範囲のセルを囲うように罫線を引く (ctrl + q)
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
' 選択セルを"塗りつぶしなし"にする (ctrl + shift + n)
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
' 選択セルを黄色に塗りつぶす (ctrl + shift + y)
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
' 結合セルの高さを自動調節 (ctrl + shift + r)
' Margin でセルの高さを調整可能
'
    Dim Margin                  As Integer      ' 自動調整時マージン[px]
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
        iBondWidth = GetWidthOfMergedCells(r)
        
        '// 結合時の各セルの高さを取得
        Call GetHightOfMergedCells(r, iArRow)
        
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
            bRet = GetAddressesOfMergedCells(r, sArCell)
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
            Call SetHightOfMergedCells(r, iBeforeHeight, iArRow)
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub


Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' 選択セルの横列に合わせてオートフィル (ctrl + shift + r)
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


Sub PasteWithoutBlankRowCells()
Attribute PasteWithoutBlankRowCells.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' 結合セルの空白部分を除いて、コピー範囲のセルをペーストする
' 列方向結合に対応、複数列のコピーは不可 (ctrl + shift + v)
'
    Dim data_obj As New DataObject      ' クリップボード参照の為のDataObject
    Dim cbFormat As Variant
    Dim trimmed_txt As String           ' 空白文字等を削除した文字列用
    Dim cells_array As Variant          ' コピーセル範囲内の値を要素とする配列
    Dim i As Long                       ' ループカウンタ
    Dim paste_index As Long             ' ペースト先セルのindex(選択セルを基準に移動)
    
    paste_index = 0
    ' クリップボードのデータがテキスト以外ならば終了
    cbFormat = Application.ClipboardFormats
    If cbFormat(1) <> 0 Then
        Exit Sub
    End If
    
    
    ' コピーセル範囲の値を配列化
    data_obj.GetFromClipboard
    ' クリップボードの文字列から空白文字等を削除し整形
    trimmed_txt = Replace(data_obj.GetText, vbTab, "")
    trimmed_txt = Replace(trimmed_txt, vbCr, "")
    trimmed_txt = Replace(trimmed_txt, vbCrLf, "")
    cells_array = Split(trimmed_txt, vbLf)

    
    ' 選択セルを基準に値をペースト
    For i = 0 To UBound(cells_array) - 1
        ' 値の入ったセルのみペースト
        If cells_array(i) <> "" Then
            ' 文字列としてセルに代入
            Selection.Offset(paste_index, 0).Value = "'" & cells_array(i)
            paste_index = paste_index + 1
        End If
    Next
    
End Sub
