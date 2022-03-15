Option Explicit

' プロジェクトディレクトリ
const Apli_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Apli/"
const Boot_path = "C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Boot/"

' pythonスクリプト
const toggle_src = "toggle_src.py"

call main


sub main()
    dim fso 
    set fso = CreateObject("Scripting.FileSystemObject")

    dim script
    script = fso.GetParentFolderName(Editor.ExpandParameter("$I")) & "/" & toggle_src
    dim wsh
    set wsh = CreateObject("WScript.Shell")

    dim command
    command = "--command=toggle"
    dim tgt_path
    tgt_path = "--tgt_path=" & replace(Editor.GetFileName, "\", "/")

    dim cl_input
    cl_input = join(array("cmd.exe /c python", script, command, tgt_path), " ")
    ' コマンドプロンプト非表示で同期実行
    dim is_err
    is_err = wsh.Run(cl_input, 0, True)
    if is_err then
        MsgBox "toggling failed"
        exit sub
    end if
    
    Editor.TagMake()

    call wsh.Run(cl_input, 0, True)
    
end sub