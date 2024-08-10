' 管理者として実行すること
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("SYSTEM")
strNewPath = "C:\Users\user\git"
strCurrentPath = objEnv("PATH")
' パスがすでに存在するかチェック
If InStr(1, strCurrentPath, strNewPath, vbTextCompare) = 0 Then
    ' 新しいパスを追加
    objEnv("PATH") = strCurrentPath & ";" & strNewPath
    WScript.Echo "新しいパスが追加されました: " & strNewPath
Else
    WScript.Echo "パスは既に存在します: " & strNewPath
End If