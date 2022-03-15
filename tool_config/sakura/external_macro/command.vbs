Option Explicit

function search_command(tgt_path, pattern)
	' TODO: RunでのファイルI/Oのオーバーヘッド確認, Execとの比較検討

	dim wsh
	set wsh = CreateObject("WScript.Shell")
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")

	dim command
	command = "--command=search"
	dim script
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	tgt_path = "--tgt_path=" & tgt_path
	pattern = "--pattern=" & pattern

	dim cl_input
	cl_input = join(array("cmd.exe /c python", script, command, tgt_path, pattern), " ")

	search_command = wsh.Exec(cl_input).StdOut.ReadLine
end function


function project_command(tgt_proj)
	dim wsh
	set wsh = CreateObject("WScript.Shell")
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")

	dim command
	command = "--command=project"
	dim script
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	dim proj_name
	proj_name = "--proj_name=" & tgt_proj

	dim cl_input
	cl_input = join(array("cmd.exe /c python", script, command, proj_name), " ")

	project_command = wsh.Exec(cl_input).StdOut.ReadLine
end function


function toggle_command(tgt_path)
	dim wsh
	set wsh = CreateObject("WScript.Shell")
	dim fso 
	set fso = CreateObject("Scripting.FileSystemObject")

	dim script
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	dim command
	command = "--command=toggle"
	tgt_path = "--tgt_path=" & tgt_path

	dim cl_input
	cl_input = join(array("cmd.exe /c python", script, command, tgt_path), " ")

	dim is_err
    ' コマンドプロンプト非表示で同期実行
    is_err = wsh.Run(cl_input, 0, True)

	toggle_command = is_err
end function