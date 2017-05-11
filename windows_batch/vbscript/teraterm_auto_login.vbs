'--------------------------
'定数定義
'--------------------------
'exeのpathを記載してください
Const exe_file = """hogehoge.exe"""

'--------------------------
'デフォルト設定
'--------------------------
Dim host, user, passwd
host   = ""
user   = ""
passwd = ""


'--------------------------
'コマンド引数のロード
'--------------------------
Dim count
count = WScript.Arguments.Count

'引き数が不足している時は強制終了
if 0 < count AND count < 3 then
    WScript.echo("too few arguments!!!")
    WScript.Quit(-1)
end if

'引き数指定があればデフォルト値を上書き
if count <> 0 then
    host   = WScript.Arguments(0)
    user   = WScript.Arguments(1)
    passwd = WScript.Arguments(2)
end if

'--------------------------
'起動オプションを定義
'--------------------------
Dim params
params = Array( _
    host, _
    " /ssh2", _ 
    "/auth=password", _
    "/user=" & user, _
    "/passwd=" & passwd, _
    "/LA=E")

'--------------------------
'起動コマンド生成
'--------------------------
Dim opt_params
For Each param In params
    opt_param = opt_param + param & " "
Next

'--------------------------
'コマンド実行
'--------------------------
Dim wshell
Set wshell = CreateObject("WScript.Shell")
wshell.run(exe_file & " " & opt_param)
Set wshell = Nothing

