Attribute VB_Name = "MyFunc"
Option Explicit

' ------------------- ���L�֐��͓K�p�������u�b�N�ɃR�s�[���Ďg�p���� ------------------------------


Function CONCAT(ParamArray par())
'
' �I��͈͂̃Z���̕��������������
' �����͐���������
' +input
'       par() : object, �Z���͈�
' +output
'       �Z���͈͓��̒l����������������

  Dim i As Integer
  Dim tR As Range
  
  CONCAT = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        CONCAT = CONCAT & tR.Value2
      Next
    Else
      CONCAT = CONCAT & par(i)
    End If
  Next
End Function


Function StrReverseHex(a_str) As String
'
' ���g���G���f�B�A���̕�������A�h���X���ɕϊ�����
' +input
'       a_str : str, ����2n�̕�����
' +output
'       �ϊ��㕶����

    Dim address_order_hex As String     ' �o�͗p�ϐ�
    Dim len_s As Long                   ' ����������̒���
    Dim i As Long                       ' ���[�v�J�E���^
    
    address_order_hex = ""
    len_s = Len(a_str)
    
    ' �����`�F�b�N
    If (len_s Mod 2) <> 0 Then
        Debug.Print "����������̒�����2�̔{���ł͂���܂���"
    End If
    
    ' �擪����2�������擾���A�A�h���X���ɕ��ёւ�
    For i = 1 To len_s Step 2
        address_order_hex = Mid(a_str, i, 2) & address_order_hex
    Next
    StrReverseHex = address_order_hex
End Function

