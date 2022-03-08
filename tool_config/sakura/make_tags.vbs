Option Explicit

' プロジェクトディレクトリ
const Apli_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Apli/"
const Boot_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Boot/"

' pythonスクリプト
const toggle_src = "toggle_src.py"

call main


sub main()
    dim tgt_dir
	if InStr(Editor.GetFileName, "Apli") then
		tgt_dir = Apli_path
	elseif InStr(Editor.GetFileName, "Boot") Then
		tgt_dir = Boot_path
	else
		MsgBox "Please open source file in registered project."
		exit sub
	end if

    dim fso 
    set fso = CreateObject("Scripting.FileSystemObject")

    dim py_script
    py_script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & toggle_src

    dim wsh
    set wsh = CreateObject("WScript.Shell")

    dim command
    command = "python " & py_script & " " & tgt_dir
    ' コマンドプロンプト非表示で同期実行
    call wsh.Run(command, 0, True)
    
    Editor.TagMake()

    call wsh.Run(command, 0, True)
    
end sub