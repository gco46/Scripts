Attribute VB_Name = "Macro"
Option Explicit


Sub delete_name_and_style()
'
' 名前定義、書式定義を全削除
'

    On Error Resume Next

    '名前定義を全削除（名前を関数その他に有効活用している場合はここは削除）

    Dim N As name
    For Each N In ActiveWorkbook.Names
        N.Delete
    Next

    '書式（スタイル）定義を全削除

    Dim M()
    Dim j As Integer
    Dim i As Integer

    j = ActiveWorkbook.Styles.Count
    ReDim M(j)
    For i = 1 To j
        M(i) = ActiveWorkbook.Styles(i).name
    Next
    For i = 1 To j
        If InStr("Hyperlink,Normal,Followed Hyperlink", _
                    M(i)) = 0 Then
            ActiveWorkbook.Styles(M(i)).Delete
        End If
    Next

End Sub


Sub ExportAll()
'
' ブックに紐づいたマクロを一括でエクスポートする
' ブックが開かれている場合：開いているブックを対象とする
' ブックが開かれていない場合：個人用マクロブックを対象とする
'
    Dim module                  As VBComponent      '// モジュール
    Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath                                       '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook                                  '// 処理対象ブックオブジェクト
    Dim OutputPath                                  '// エクスポート先パス
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            OutputPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    '// ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// ブックが開かれている場合は表示しているブックを対象とする
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.Path
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        '// クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        '// 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// その他
        Else
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        '// エクスポート実施
        sFilePath = OutputPath & "\" & module.name & "." & extension
        Call module.Export(sFilePath)
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub
