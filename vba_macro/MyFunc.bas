Attribute VB_Name = "MyFunc"
Option Explicit

' ------------------- 下記関数は適用したいブックにコピーして使用する ------------------------------


Function CONCAT(ParamArray par())
'
' 選択範囲のセルの文字列を結合する
' 走査は水平→垂直
' +input
'       par() : object, セル範囲
' +output
'       セル範囲内の値を結合した文字列

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
' リトルエンディアンの文字列をアドレス順に変換する
' +input
'       a_str : str, 長さ2nの文字列
' +output
'       変換後文字列

    Dim address_order_hex As String     ' 出力用変数
    Dim len_s As Long                   ' 引数文字列の長さ
    Dim i As Long                       ' ループカウンタ
    
    address_order_hex = ""
    len_s = Len(a_str)
    
    ' 引数チェック
    If (len_s Mod 2) <> 0 Then
        Debug.Print "引数文字列の長さが2の倍数ではありません"
    End If
    
    ' 先頭から2文字ずつ取得し、アドレス順に並び替え
    For i = 1 To len_s Step 2
        address_order_hex = Mid(a_str, i, 2) & address_order_hex
    Next
    StrReverseHex = address_order_hex
End Function

