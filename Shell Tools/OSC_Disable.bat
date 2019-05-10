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
