Attribute VB_Name = "Macro"
Option Explicit


Sub delete_name_and_style()
'
' ���O��`�A������`��S�폜
'

    On Error Resume Next

    '���O��`��S�폜�i���O���֐����̑��ɗL�����p���Ă���ꍇ�͂����͍폜�j

    Dim N As name
    For Each N In ActiveWorkbook.Names
        N.Delete
    Next

    '�����i�X�^�C���j��`��S�폜

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

