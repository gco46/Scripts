Call Main

Const APLI_SCH_PATH = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Apli/PJ/SRV/SCH/Sch.c"

Sub Main()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If (fso.FileExists(APLI_SCH_PATH)) Then
		Editor.FileOpen (APLI_SCH_PATH)
	Else
		MsgBox "Cannot open Sch.c in Application."
	End If
End Sub