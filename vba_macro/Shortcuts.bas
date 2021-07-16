Attribute VB_Name = "Shortcuts"
Option Explicit

Private Sub startMacro()
'
' �}�N�����s�O����(�������̂��߂Ɋe�폈���𖳌���)
'
    With Application
        .ScreenUpdating = False              '�`����ȗ�
        .Calculation = xlCalculationManual   '�蓮�v�Z
        .DisplayAlerts = False               '�x�����ȗ�
        .EnableEvents = False                '�C�x���g����
    End With
End Sub

Private Sub endMacro()
'
' �}�N�����s�㏈��(�e�폈���L����)
'
    With Application
        .ScreenUpdating = True                '�`�悷��
        .Calculation = xlCalculationAutomatic '�����v�Z
        .DisplayAlerts = True                 '�x�����s��
        .EnableEvents = True                  '�C�x���g�L��
    End With
End Sub


Private Sub executeMacro(method As String)
'
' �}�N�����s�֐�
' method �ŃR�[���o�b�N�֐����w�肷��
'
    Dim cbObj As New Callback
    startMacro
    executeCallback Array(cbObj, method)
    endMacro
End Sub

Private Sub executeCallback(cb_arr)
    CallByName cb_arr(0), cb_arr(1), VbMethod
End Sub
 
 
Sub ShowMultiBook()
Attribute ShowMultiBook.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' �J���Ă���u�b�N�����ɕ��ׂĕ\��
'
    executeMacro "show_multi_book"
End Sub


Sub ShowOneBook()
Attribute ShowOneBook.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' �A�N�e�B�u�ȃu�b�N��S��ʕ\��
'
    executeMacro "show_one_book"
End Sub


Sub NoPaint()
Attribute NoPaint.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' �I���Z����"�h��Ԃ��Ȃ�"�ɂ��� (ctrl + shift + n)
'
    executeMacro "paint_blank"
End Sub


Sub YellowPaint()
Attribute YellowPaint.VB_ProcData.VB_Invoke_Func = "Y\n14"
'
' �I���Z�������F�ɓh��Ԃ� (ctrl + shift + y)
'
    executeMacro "paint_yellow"
End Sub


Sub DrawLatticeLine()
Attribute DrawLatticeLine.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �I���Z���Ɋi�q��̌r���������ictrl + q�j
'
    executeMacro "draw_lattice_line"
End Sub


Sub AutoFitMergedCellsHeight()
Attribute AutoFitMergedCellsHeight.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' �����Z���̍������������� (ctrl + shift + r)
'
    executeMacro "fit_merged_cells_height"
End Sub


Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' �I���Z���̉���ɍ��킹�ăI�[�g�t�B�� (ctrl + shift + f)
'
    executeMacro "auto_fill"
End Sub


Sub PasteWithoutBlankRowCells()
Attribute PasteWithoutBlankRowCells.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' �����Z���̋󔒕����������āA�R�s�[�͈͂̃Z�����y�[�X�g����
' ����������ɑΉ��A������̃R�s�[�͕s�� (ctrl + shift + v)
'
    executeMacro "paste_without_blank_row_cells"
End Sub

Sub ClearInsideBorders()
Attribute ClearInsideBorders.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' �I���Z���̓����r���̂݃N���A���� (ctrl + shift + q)
'

    executeMacro "clear_inside_border"
End Sub

Sub AlignAndDistributeV()
Attribute AlignAndDistributeV.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' �I�𒆂̃V�F�C�v�����E��������&�㉺�ɐ��� (ctrl + shift + a)
'

    executeMacro "aligne_and_distribute_v"
End Sub

Sub AlignAndDistributeH()
Attribute AlignAndDistributeH.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' �I�𒆂̃V�F�C�v���㉺��������&���E�ɐ��� (ctrl + shift + s)
'

    executeMacro "aligne_and_distribute_h"
End Sub

Sub ToggleShapeGroup()
Attribute ToggleShapeGroup.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' �I�𒆂̃V�F�C�v���O���[�v��/�O���[�v�������� (ctrl + shift + g)
'

    executeMacro "toggle_shape_group"
End Sub
