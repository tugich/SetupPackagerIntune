@echo off
powershell.exe -NoProfile -NoLogo -ExecutionPolicy ByPass -File "%~dp0Upload-Intune.ps1" -IntuneWinFile "%~1" -DisplayName "%~2"
