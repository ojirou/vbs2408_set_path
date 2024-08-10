@echo off
:: 管理者として実行するよう求める
if not "%1"=="am_admin" (
    powershell -Command "Start-Process '%~dpnx0' -ArgumentList 'am_admin' -Verb runAs"
    exit /b
)
:: 以下に実行したいVBSスクリプトのパスを入力
cscript //nologo "C:\Users\user\git\vbs\set_path.vbs"