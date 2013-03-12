@ECHO OFF

:Program
CLS
ECHO E. Explorer++
ECHO D. Device Manager
ECHO C. Command Prompt
ECHO R. Rename Computer
ECHO S. Shutdown Computer (Release/Flush/Shutdown)
ECHO B. Restart Computer (Release/Flush/Restart)
ECHO I. Internet Explorer (Install ActiveX)
ECHO P. System Properties (Rename/Env Var)
ECHO M. Computer Management (EventVwr/User Groups/DiskMgmt)
ECHO A. Add/Remove Programs
ECHO Q. Quit Program
SET Program=
SET /P Program=Would you like to Rename/Shutdown/Restart the computer?: 
cscript.exe "\\jc1wsalt03\Library\Packages\Dantools\WriteLog.vbs" %Program%
IF NOT '%Program%'=='' SET Program=%Program:~0,1%
ECHO.
IF /I '%Program%'=='E' GOTO ProgramE
IF /I '%Program%'=='D' GOTO ProgramD
IF /I '%Program%'=='C' GOTO ProgramC
IF /I '%Program%'=='R' GOTO ProgramR
IF /I '%Program%'=='S' GOTO ProgramS
IF /I '%Program%'=='B' GOTO ProgramB
IF /I '%Program%'=='I' GOTO ProgramI
IF /I '%Program%'=='P' GOTO ProgramP
IF /I '%Program%'=='M' GOTO ProgramM
IF /I '%Program%'=='A' GOTO ProgramA
IF /I '%Program%'=='Q' GOTO ProgramQ

ECHO.
GOTO Program
:ProgramE
GOTO Explorer
:ProgramD
GOTO DeviceManager
:ProgramC
GOTO cmdLOOP
:ProgramR
GOTO RenameLoop
:ProgramS
GOTO ShutdownLoop
:ProgramB
GOTO RestartLoop
:ProgramI
GOTO InternetExplorer
:ProgramP
GOTO SysProp
:ProgramM
GOTO CompMgmt
:ProgramA
GOTO Appwiz
:ProgramQ
GOTO :End
:ProgramAgain
PING -n 1 127.0.0.1 >NUL
ECHO.
ECHO.

:Explorer
CLS
ECHO Y. Open Explorer++
ECHO N. Go To Main Menu
set /p Explorer=Would you like to open Explorer++? (Y/N) : 
IF NOT '%Explorer%'=='' SET Explorer=%Explorer:~0,1%
ECHO.
IF /I '%Explorer%'=='Y' GOTO ExplorerYES
IF /I '%Explorer%'=='N' GOTO ExplorerNO
ECHO.
GOTO ExplorerLoop
:ExplorerYES
ECHO Opening Explorer++.
"\\jc1dfs2\applications\desktop\copyBat\Explorer++.exe"
GOTO ExplorerAgain
:ExplorerNO
ECHO Canceling Explorer++.
GOTO ExplorerAgain
:ExplorerAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:Appwiz
CLS
ECHO Y. Open Add/Remove Programs
ECHO N. Go To Main Menu
set /p AppWiz=Would you like to open Add/Remove Programs? (Y/N) : 
IF NOT '%AppWiz%'=='' SET AppWiz=%AppWiz:~0,1%
ECHO.
IF /I '%AppWiz%'=='Y' GOTO AppWizYES
IF /I '%AppWiz%'=='N' GOTO AppWizNO
ECHO.
GOTO AddRemoveLoop
:AppWizYES
ECHO Opening Add/Remove Programs.
start /wait appwiz.cpl
GOTO AppWizAgain
:AppWizNO
ECHO Canceling Add/Remove Programs.
GOTO AppWizAgain
:AppWizAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:DeviceManager
CLS
ECHO Y. Open Device Manager
ECHO N. Go To Main Menu
set /p DevMgr=Would you like to open Device Manager? (Y/N) : 
IF NOT '%DevMgr%'=='' SET DevMgr=%DevMgr:~0,1%
ECHO.
IF /I '%DevMgr%'=='Y' GOTO DevMgrYES
IF /I '%DevMgr%'=='N' GOTO DevMgrNO
ECHO.
GOTO DeviceManagerLoop
:DevMgrYES
ECHO Opening Device Manager.
start /wait devmgmt.msc
GOTO DevMgrAgain
:DevMgrNO
ECHO Canceling Device Manager.
GOTO DevMgrAgain
:DevMgrAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:CompMgmt
CLS
ECHO Y. Open Computer Management
ECHO N. Go To Main Menu
set /p CompMgmt=Would you like to open Computer Management? (Y/N) : 
IF NOT '%CompMgmt%'=='' SET CompMgmt=%CompMgmt:~0,1%
ECHO.
IF /I '%CompMgmt%'=='Y' GOTO CompMgmtYES
IF /I '%CompMgmt%'=='N' GOTO CompMgmtNO
ECHO.
GOTO CompMgmtLoop
:CompMgmtYES
ECHO Opening Computer Management.
start /wait compmgmt.msc
GOTO CompMgmtAgain
:CompMgmtNO
ECHO Canceling Computer Management.
GOTO CompMgmtAgain
:CompMgmtAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:SysProp
CLS
ECHO Y. Open System Properties
ECHO N. Go To Main Menu
set /p SysProperties=Would you like to open System Properties? (Y/N) : 
IF NOT '%SysProperties%'=='' SET SysProperties=%SysProperties:~0,1%
ECHO.
IF /I '%SysProperties%'=='Y' GOTO SysPropYES
IF /I '%SysProperties%'=='N' GOTO SysPropNO
ECHO.
GOTO SysPropertiesLoop
:SysPropYES
ECHO Opening System Properties.
start /wait sysdm.cpl
GOTO SysPropAgain
:SysPropNO
ECHO Canceling System Properties.
GOTO SysPropAgain
:SysPropAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:InternetExplorer
CLS
ECHO Y. Open Internet Explorer
ECHO N. Go To Main Menu
set /p IExplore=Would you like to open Internet Explorer? (Y/N) : 
IF NOT '%IExplore%'=='' SET IExplore=%IExplore:~0,1%
ECHO.
IF /I '%IExplore%'=='Y' GOTO IExploreYes
IF /I '%IExplore%'=='N' GOTO IExploreNo
ECHO.
GOTO InternetExplorerLoop
:IExploreYES
ECHO Opening Internet Explorer.
"c:\program files\internet explorer\iexplore.exe"
GOTO IExploreAgain
:IExploreNO
ECHO Canceling Internet Explorer.
GOTO IExploreAgain
:IExploreAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:cmdLOOP
CLS
ECHO Y. Open Command Prompt?
ECHO N. Do Not Open Command Prompt?
SET cmd=
SET /P cmd=Would you like to Open Command Prompt? (Y/N): 
IF NOT '%cmd%'=='' SET Restart=%cmd:~0,1%
ECHO.
IF /I '%cmd%'=='Y' GOTO cmdYes
IF /I '%cmd%'=='N' GOTO cmdNo
ECHO.
GOTO cmdLoop
:cmdYes
ECHO Opening Command Prompt.
start cmd.exe
GOTO cmdAgain
:cmdNo
ECHO No Command Prompt.
GOTO cmdAgain
:cmdAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:RenameLoop
CLS
ECHO Y. Rename Computer
ECHO N. Do Not Rename Computer
SET Rename=
SET /P Rename=Would you like to rename the computer? (Y/N): 
IF NOT '%Rename%'=='' SET Rename=%Rename:~0,1%
ECHO.
IF /I '%Rename%'=='Y' GOTO RenameYES
IF /I '%Rename%'=='N' GOTO RenameNO
ECHO.
GOTO RenameLoop
:RenameYES
ECHO Renaming Computer.
start /wait cscript.exe \\jc1dfs2\applications\desktop\CopyBat\Rename.vbs
GOTO RenameAgain
:RenameNO
ECHO Canceling Rename.
GOTO RenameAgain
:RenameAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:ShutdownLOOP
CLS
ECHO Y. Shutdown Computer
ECHO N. Do Not Shutdown Computer
ECHO T. Shutdown Without Release/Flush
SET Shutdown=
SET /P Shutdown=Would you like to shutdown the computer? (Y/N): 
IF NOT '%Shutdown%'=='' SET Shutdown=%Shutdown:~0,1%
ECHO.
IF /I '%Shutdown%'=='Y' GOTO ShutdownYes
IF /I '%Shutdown%'=='T' GOTO ShutdownNoFlush
IF /I '%Shutdown%'=='N' GOTO ShutdownNo
ECHO.
GOTO ShutdownLoop
:ShutdownYes
ECHO Shutting Down Computer.
del c:\temp\shutdown.bat
copy \\jc1dfs2\applications\desktop\CopyBat\shutdown.bat c:\temp\shutdown.bat
echo.
C:\temp\shutdown.bat
GOTO ShutdownAgain
:ShutdownNoFlush
ECHO Shutting Down Computer.
shutdown -s -t 2 /f
echo.
GOTO ShutdownAgain
:ShutdownNo
ECHO Canceling Shutdown.
GOTO ShutdownAgain
:ShutdownAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.

:RestartLOOP
CLS
ECHO Y. Restart Computer
ECHO N. Do Not Restart Computer
ECHO T. Restart Without Release/Flush
SET Restart=
SET /P Restart=Would you like to Restart the computer? (Y/N): 
IF NOT '%Restart%'=='' SET Restart=%Restart:~0,1%
ECHO.
IF /I '%Restart%'=='Y' GOTO RestartYes
IF /I '%Restart%'=='T' GOTO RestartNoFlush
IF /I '%Restart%'=='N' GOTO RestartNo
ECHO.
GOTO RestartLoop
:RestartYes
ECHO Restarting Computer.
del c:\temp\restart.bat
copy \\jc1dfs2\applications\desktop\CopyBat\restart.bat c:\temp\restart.bat
echo.
c:\temp\restart.bat
GOTO RestartAgain
:RestartNoFlush
ECHO Restarting Computer.
shutdown -r -t 2 /f
echo.
GOTO RestartAgain
:RestartNo
ECHO Canceling Computer Restart.
GOTO RestartAgain
:RestartAgain
PING -n 1 127.0.0.1 >NUL
GOTO Program
ECHO.
ECHO.
:End
Pause