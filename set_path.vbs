' �Ǘ��҂Ƃ��Ď��s���邱��
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("SYSTEM")
strNewPath = "C:\Users\user\git"
strCurrentPath = objEnv("PATH")
' �p�X�����łɑ��݂��邩�`�F�b�N
If InStr(1, strCurrentPath, strNewPath, vbTextCompare) = 0 Then
    ' �V�����p�X��ǉ�
    objEnv("PATH") = strCurrentPath & ";" & strNewPath
    WScript.Echo "�V�����p�X���ǉ�����܂���: " & strNewPath
Else
    WScript.Echo "�p�X�͊��ɑ��݂��܂�: " & strNewPath
End If