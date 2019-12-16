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
    
    '// ˆø”•¶š—ñ‚Ì•¶š—ñ’·‚ğæ“¾
    iLen = Len(a_str)
    
    '// •¶š—ñ’·‚ª‹ô”‚Å‚È‚¢ê‡
    If (iLen Mod 2) <> 0 Then
        Debug.Print "‚Q•¶š‚¸‚Â‚Ìˆø”•¶š—ñ‚Å‚È‚¢"
        Exit Function
    End If
    
    '// ˆø”•¶š—ñ‚ğ”½“]
    sRev = StrReverse(a_str)
    
    '// –ß‚è’l•¶š—ñ‚ğ‰Šú‰»
    sRet = ""
    
    '// ‚Q•¶š‚¸‚Âƒ‹[ƒv
    For i = 1 To iLen Step 2
        '// ‚Q•¶š–Ú‚ğæ‚Éæ“¾
        sRet = sRet & Mid(sRev, i + 1, 1)
        '// ‚P•¶š–Ú‚ğŒã‚Éæ“¾
        sRet = sRet & Mid(sRev, i, 1)
    Next
    
    StrReverseHex = sRet
End Function
