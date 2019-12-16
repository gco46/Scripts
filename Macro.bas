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

