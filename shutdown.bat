c:
cd c:\temp
ipconfig/release
ipconfig/flushdns
taskkill /f /IM perl.exe
@echo
c:\WINDOWS\system32\shutdown.exe -s -t 5 -c "Computer is shutting down"
pause