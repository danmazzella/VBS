' Retrieve the first argument (index 0).  
strSoundFile = Wscript.Arguments(0)  

' Retrieve the second argument.  
playVol = Wscript.Arguments(1)  

' Retrieve the third argument.  
prevVol = Wscript.Arguments(2)  

Set objShell = CreateObject("Wscript.Shell")

strVolume = "c:\temp\sound\nircmdc mutesysvolume 0 master 0"
objshell.run strVolume, 0, True

strVolume = "c:\temp\sound\nircmdc setsysvolume " & playVol & " master 0"
objshell.run strVolume, 0, True

strCommand = "c:\temp\sound\sndrec32.exe /play /close " & "c:\temp\sound\" & strSoundFile
objShell.Run strCommand, 0, True 


strVolume = "c:\temp\sound\nircmdc setsysvolume " & prevVol & " master 0"
objshell.run strVolume, 0, True 