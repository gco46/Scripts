Call Main

const BOOT_SCH_PATH = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Boot/Bootloader/PJ/SRV/SCH/Sch.c"

Sub Main()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	If (fso.FileExists(BOOT_SCH_PATH)) Then
		Editor.FileOpen (BOOT_SCH_PATH)
	Else
		MsgBox "Cannot open Sch.c in Bootloader."
	End If
End Sub