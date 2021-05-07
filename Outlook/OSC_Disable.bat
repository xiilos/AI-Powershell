reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

reg query HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\AddIns\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\AddIns\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins\OscAddin.Connect

if %ERRORLEVEL% EQU 0 (
reg add HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins\OscAddin.Connect /t REG_DWORD /v LoadBehavior /d 0 /f
)

#Renaming .dll to .backup

Rename  "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALCONNECTOR.DLL" SOCIALCONNECTOR.Backup
Rename  "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALPROVIDER.DLL" SOCIALPROVIDER.Backup

if %ERRORLEVEL% EQU 1 (
Rename  "C:\Program Files\Microsoft Office\root\Office16\SOCIALCONNECTOR.DLL" SOCIALCONNECTOR.Backup
Rename  "C:\Program Files\Microsoft Office\root\Office16\SOCIALPROVIDER.DLL" SOCIALPROVIDER.Backup
)

Rename  "C:\Program Files (x86)\Microsoft Office\root\Office15\SOCIALCONNECTOR.DLL" SOCIALCONNECTOR.Backup
Rename  "C:\Program Files (x86)\Microsoft Office\root\Office15\SOCIALPROVIDER.DLL" SOCIALPROVIDER.Backup

if %ERRORLEVEL% EQU 1 (
Rename  "C:\Program Files\Microsoft Office\root\Office15\SOCIALCONNECTOR.DLL" SOCIALCONNECTOR.Backup
Rename  "C:\Program Files\Microsoft Office\root\Office15\SOCIALPROVIDER.DLL" SOCIALPROVIDER.Backup
)