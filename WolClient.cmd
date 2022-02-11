@echo off
pushd %~dp0
c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Executionpolicy Bypass -File "%~n0.ps1"
popd