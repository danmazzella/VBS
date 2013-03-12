c:
cd c:\
ipconfig/release
ipconfig/flushdns
taskkill /f /IM perl.exe
@echo
c:\WINDOWS\system32\shutdown.exe -r -t 5 -c "Computer is restarting"
pause