Attribute VB_Name = "MyFunc"
Option Explicit


Function CONCAT(ParamArray par())
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
    Dim sRev    As String
    Dim i       As Long
    Dim iLen    As Long
    Dim sRet    As String
    
    '// ����������̕����񒷂��擾
    iLen = Len(a_str)
    
    '// �����񒷂������łȂ��ꍇ
    If (iLen Mod 2) <> 0 Then
        Debug.Print "�Q�������̈���������łȂ�"
        Exit Function
    End If
    
    '// ����������𔽓]
    sRev = StrReverse(a_str)
    
    '// �߂�l�������������
    sRet = ""
    
    '// �Q���������[�v
    For i = 1 To iLen Step 2
        '// �Q�����ڂ��Ɏ擾
        sRet = sRet & Mid(sRev, i + 1, 1)
        '// �P�����ڂ���Ɏ擾
        sRet = sRet & Mid(sRev, i, 1)
    Next
    
    StrReverseHex = sRet
End Function
