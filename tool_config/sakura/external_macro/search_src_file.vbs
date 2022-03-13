Option Explicit

' プロジェクトディレクトリ
const Apli_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Apli/"
const Boot_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Boot/"
const os_path = "D:/project/git_root/h8-embedded-os/07/os"

' 対象の拡張子
dim extension_list
extension_list = Array("c", "h", "s", "asm", "800")

' call main
call main2

sub main2()
	' TODO: json設定読込用
	dim tgt_file_name
	tgt_file_name = InputBox("Input file name:")
	if tgt_file_name = "" then
		exit sub
	end if

	tgt_file_name = LCase(tgt_file_name)

	dim wsh
    set wsh = CreateObject("WScript.Shell")

	dim command
	command = "--command=search"
	dim tgt_path
	tgt_path = "--tgt_path=" & os_path
	dim pattern
	pattern = "--pattern=" & tgt_file_name
	dim script
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")
	script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & "toggle_src.py"
	dim tmp_txt
	tmp_txt = "tmp.txt"

	dim cl_input
	' TODO: RunでのファイルI/Oのオーバーヘッド確認, Execとの比較検討
	' cl_input = join(array("cmd.exe /c python", script, command, tgt_path, pattern, ">", tmp_txt), " ")
	' call wsh.Run(cl_input, 0, True)

	cl_input = join(array("cmd.exe /c python", script, command, tgt_path, pattern), " ")
	dim file_path
	file_path = wsh.Exec(cl_input).StdOut.ReadLine
	if file_path <> "" then
		Editor.FileOpen(file_path)
	else
		MsgBox "No file was found."
	end if
end sub

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

	dim tgt_file_name
	tgt_file_name = InputBox("Input file name:")
	
	if tgt_file_name = "" then
		exit sub
	end if

	tgt_file_name = LCase(tgt_file_name)

	dim fso, folder, file
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(tgt_dir)

	dim file_path
	file_path = search_file_recursive(folder, tgt_file_name)

	if file_path <> "" Then
		Editor.FileOpen(file_path)
	else
		MsgBox "No file was found."
	end if

end sub

function search_file_recursive(folder, search_ptn)
	dim file, subfolder
	dim result
	for each subfolder in folder.subFolders
		result = search_file_recursive(subfolder, search_ptn)
		if result <> "" then
			search_file_recursive = result
			exit function
		end if
	next

	for each file in folder.Files
		if match_srcfile(file.Name, search_ptn) then
			search_file_recursive = file.Path
			exit function
		end if
	next

	search_file_recursive = ""
end function

function match_srcfile(file_name, pattern)
	dim fso
	set fso = CreateObject("Scripting.FileSystemObject")
	dim extension
	extension = fso.GetExtensionName(file_name)
	
	dim result
	result = false
	
	dim is_src_file
	' Filter()で配列中要素の位置を取得
	' 存在しなければ-1が返ることを利用し対象拡張子のファイルかどうか判定

	is_src_file = is_tgt_file(extension)

	if is_src_file then
		' 大文字小文字を区別しない(全て小文字で比較)
		file_name = LCase(file_name)
		
		result = startswith(file_name, pattern)
	end if
	
	match_srcfile = result

end function

Function startswith(tgt_str, prefix)
    startswith = False
    If Len(tgt_str) < Len(prefix) Then
        Exit Function
    End If
    If InStr(1, tgt_str, prefix) = 1 Then
        startswith = True
    End If
End Function


function is_tgt_file(ext)
	dim tgt_ext
	dim result
	for each tgt_ext in extension_list
		result = strComp(tgt_ext, ext, 1)
		if result = 0 then
			is_tgt_file = true
			exit function
		end if
	next

	is_tgt_file = false
end function