@echo off

setlocal enabledelayedexpansion
echo.

net session >nul 2>&1
if NOT %errorLevel% == 0 (
     echo You do not have sufficient rights to execute this command
     echo Please execute as Administrator
     exit /B 1
)

echo Testing for Click2Run
rem Check 32 on 32 or 64 on 64
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} 2>NUL >NUL
IF %ERRORLEVEL%==0 (
	echo Click2Run Detected, Checking if patch applied
	reg query HKLM\SOFTWARE\Classes\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} 2>NUL >NUL
	IF !ERRORLEVEL!==1 (
		echo Unpatched Click2Run Detected, patching
		rem IConverterSession
		reg copy HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} HKLM\SOFTWARE\Classes\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} /s /f
		rem IMimeMessage
		reg copy HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} HKLM\SOFTWARE\Classes\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} /s /f
	) ELSE (
		echo Click2Run is already patched, nothing to do
	)
	goto end
)

rem Check 32 on 64
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} 2>NUL >NUL
IF %ERRORLEVEL%==0 (
	echo Click2Run 32on64 Detected, Checking if patch applied
	reg query HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} 2>NUL >NUL
	IF !ERRORLEVEL!==1 (
		echo Unpatched Click2Run Detected, patching
		rem IConverterSession
		reg copy HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{4E3A7680-B77A-11D0-9DA5-00C04FD65685} /s /f
		rem IMimeMessage
		reg copy HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\{9EADBD1A-447B-4240-A9DD-73FE7C53A981} /s /f
	) ELSE (
		echo Click2Run is already patched, nothing to do
	)
	goto end
)

echo No Click2Run, nothing to do
echo.
:end