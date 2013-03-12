on error resume next
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run """\\jc1wsalt03\Library\Packages\Dantools\form.hta""", 1
