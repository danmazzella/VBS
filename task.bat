@echo off
setlocal enabledelayedexpansion
for /f "skip=4 delims=" %%A in (c:\temp\temp.txt) do (
    set /a current+=1
    echo %%A > !current!
    set memuse=%%A
    set memuse=!memuse:~-12!
    set memuse=!memuse: =!
    set memuse=!memuse:,=!
    set memuse=!memuse:k=!
echo Is !current! GTR 1?
rem pause > nul
    if !current! gtr 1 (
echo     it is!
    set using=!current!
        for /l %%F in (!current!,-1,1) do (
echo     On file: %%F
rem pause > nul
            for /f "delims=" %%G in (%%F) do (
                set Omemuse=%%G
                set Omemuse=!Omemuse:~-12!
                set Omemuse=!Omemuse: =!
                set Omemuse=!Omemuse:,=!
                set Omemuse=!Omemuse:k=!
echo     Is !memuse! GTR !Omemuse!?
rem pause > nul
                If !memuse! gtr !Omemuse! (
echo         It is
rem pause > nul
echo         Renaming %%F to TEMPORARY
        Ren %%F TEMPORARY
rem pause > nul
echo         Renaming !using! to %%F
        Ren !using! %%F
rem pause > nul
echo         Renaming TEMPORARY to !using!
        Ren TEMPORARY !using!
rem pause > nul
echo        Changing !using! to %%F
        Set using=%%F
rem pause > nul
                )
            )
        )
    )
)
rem pause > nul
cls
set file=1
:oddloop
    if not exist %file% goto complete
    set /p info=<%file%
    set procname=%info:~0,28%
    set procname=%procname: =%
    set mem=%info:~-12%
    set mem=%mem: =%
rem echo SET # = %procname%
    set #=%procname%
    set length=0
    set length2=0
    :loop
    IF defined # (SET #=%#:~1%&SET /A length += 1&goto loop)
rem echo Set #2 - %mem%
    set #2=%mem%
    :loop2
    IF defined #2 (SET #2=%#2:~1%&SET /A length2 += 1&goto loop2)
rem echo Spacecount = 36 - %length% - %length2%
    set /a spacecount=36-%length%-%length2%
    for /l %%Q in (1,1,%spacecount%) do (
        set procname=!procname! 
    )
    echo %procname%%mem% >> c:\temp\tasks.txt
    Del %file%
    set /a file+=1
goto oddloop
:complete
Del c:\temp\Temp.txt