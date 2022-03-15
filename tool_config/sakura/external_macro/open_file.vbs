Option Explicit
' dim fso
' set fso = CreateObject("Scripting.FileSystemObject")
' dim vbs_lib
' vbs_lib = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "search_src_file.vbs"

' Include(vbs_lib)

call main()

sub main()
	dim wsh
    set wsh = CreateObject("WScript.Shell")
	dim command
	command = "--command=project"

	dim script
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	dim cl_input
	cl_input = join(array("cmd.exe /c python", script, command), " ")

	dim proj_csv
	proj_csv = wsh.Exec(cl_input).StdOut.ReadLine
	
	dim message
	message = replace(proj_csv, ",", vbcr & "  ")
	message = "Choose project:" & vbcr & "  " & message
	
	dim tgt_proj
	tgt_proj = InputBox(message)
	if tgt_proj = "" then
		exit sub
	end if

	dim proj_name
	proj_name = "--proj_name=" & tgt_proj
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	cl_input = join(array("cmd.exe /c python", script, command, proj_name), " ")
	dim file_path
	file_path = wsh.Exec(cl_input).StdOut.ReadLine
	if file_path <> "" then
		Editor.FileOpen(file_path)
	else
		MsgBox "No file was found."
	end if
	' dim tgt_file_name
	' tgt_file_name = InputBox("Input file name:")
	' if tgt_file_name = "" then
	' 	exit sub
	' end if

	' dim file_path
	' file_path = search_src_file(tgt_file_name)

	' if file_path <> "" then
	' 	Editor.FileOpen(file_path)
	' else
	' 	MsgBox "No file was found."
	' end if
end sub

Function Include(strFile)
	'strFile：読み込むvbsファイルパス
 
	Dim objFso, objWsh, strPath
	Set objFso = CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	Set objWsh = objFso.OpenTextFile(strFile)
	ExecuteGlobal objWsh.ReadAll()
	objWsh.Close
 
	Set objWsh = Nothing
	Set objFso = Nothing
 
End Function