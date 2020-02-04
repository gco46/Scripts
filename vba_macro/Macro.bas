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


Sub ExportAll()
'
' �u�b�N�ɕR�Â����}�N�����ꊇ�ŃG�N�X�|�[�g����
' �u�b�N���J����Ă���ꍇ�F�J���Ă���u�b�N��ΏۂƂ���
' �u�b�N���J����Ă��Ȃ��ꍇ�F�l�p�}�N���u�b�N��ΏۂƂ���
'
    Dim module                  As VBComponent      '// ���W���[��
    Dim moduleList              As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                                   '// ���W���[���̊g���q
    Dim sPath                                       '// �����Ώۃu�b�N�̃p�X
    Dim sFilePath                                   '// �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook                                  '// �����Ώۃu�b�N�I�u�W�F�N�g
    Dim OutputPath                                  '// �G�N�X�|�[�g��p�X
    
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            OutputPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͌l�p�}�N���u�b�N�ipersonal.xlsb�j��ΏۂƂ���
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// �u�b�N���J����Ă���ꍇ�͕\�����Ă���u�b�N��ΏۂƂ���
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.Path
    
    '// �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        '// �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        '// �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// ���̑�
        Else
            '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
        
        '// �G�N�X�|�[�g���{
        sFilePath = OutputPath & "\" & module.name & "." & extension
        Call module.Export(sFilePath)
        
        '// �o�͐�m�F�p���O�o��
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub


Sub PasteWithoutBlankCells()
'
' �����Z���̋󔒕����������āA�R�s�[�͈͂̃Z�����y�[�X�g����
' �P���̃y�[�X�g�̂ݑΉ�
'
    Dim data_obj As New DataObject      ' �N���b�v�{�[�h�Q�Ƃׂ̈�DataObject
    Dim cbFormat As Variant
    Dim cells_array As Variant          ' �R�s�[�Z���͈͓��̒l��v�f�Ƃ���z��
    Dim i As Long                       ' ���[�v�J�E���^
    Dim paste_index As Long             ' �y�[�X�g��Z����index(�I���Z������Ɉړ�)
    
    paste_index = 0
    ' �N���b�v�{�[�h�̃f�[�^���e�L�X�g�ȊO�Ȃ�ΏI��
    cbFormat = Application.ClipboardFormats
    If cbFormat(1) <> 0 Then
        Exit Sub
    End If
    
    
    ' �R�s�[�Z���͈͂̒l��z��
    data_obj.GetFromClipboard
    cells_array = Split(data_obj.GetText, vbCrLf)
    
    ' �I���Z������ɒl���y�[�X�g
    For i = 0 To UBound(cells_array)
        ' �l�̓������Z���̂݃y�[�X�g
        If cells_array(i) <> "" Then
            Selection.Offset(paste_index, 0).Value = cells_array(i)
            paste_index = paste_index + 1
        End If
    Next
    
End Sub
