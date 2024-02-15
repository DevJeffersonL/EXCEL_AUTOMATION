@echo off
taskkill /F /IM wscript.exe /T
taskkill /F /IM cscript.exe /T
start saveOpen.vbs
