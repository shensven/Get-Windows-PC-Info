@echo off

@REM Encode: UTF-8
@REM Auther: Sven
@REM Website: https://github.com/shensven/Get-Windows-PC-Info

echo Waiting...

set myTitle=Get Windows PC Info
set outputFile=%ComputerName%.csv
set firstDelimiter=,
set secondDelimiter=,
title %myTitle%

if exist %outputFile% ( del %outputFile% )

REM IP
for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %ComputerName% ^| findstr [') do set networkIP=%%a
set "networkIP=%networkIP: =%"
echo Network IPv4%firstDelimiter%%networkIP%%secondDelimiter% >> %outputFile%

REM MAC
REM for /f "delims= tokens=1" %%a in ('getmac /FO table /NH') do set macAddress=%%a
for /f %%a in ('powershell "Get-NetAdapter | format-wide MacAddress"') do set macAddress=%%a
set "macAddress=%macAddress: =%"
echo MAC Address%firstDelimiter%%macAddress%%secondDelimiter% >> %outputFile%

REM Motherboader
for /f "delims=^= tokens=2" %%a in ('wmic baseboard get Manufacturer /format:list ^| findstr ^=') do set motherboardManufacturer=%%a
set "motherboardManufacturer=%motherboardManufacturer: =%"
for /f "delims=^= tokens=2" %%a in ('wmic baseboard get Product /format:list ^| findstr ^=') do set motherboardProduct=%%a
set "motherboardProduct=%motherboardProduct: =%"

if "%motherboardManufacturer%"=="ASUSTeKCOMPUTERINC." (
    set "motherboardManufacturer=ASUS"
) else if "%motherboardManufacturer%"=="GigabyteTechnologyCo.,Ltd." (
    set "motherboardManufacturer=Gigabyte"
) else if "%motherboardManufacturer%"=="DellInc." (
    set "motherboardManufacturer=Dell"
) else if "%motherboardManufacturer%"=="Ondatechnologycorporation" (
    set "motherboardManufacturer=ONDA"
) else if "%motherboardManufacturer%"=="OndaTechnologyCorporation" (
    set "motherboardManufacturer=ONDA"
) else if "%motherboardManufacturer%"=="UNIKATechnologyCorporation" (
    set "motherboardManufacturer=UNIKA"
)
echo Motherboard%firstDelimiter%%motherboardManufacturer% %motherboardProduct%%secondDelimiter% >> %outputFile%

REM CPU
for /f "delims=^= tokens=2" %%a in ('wmic cpu get Name /format:list ^| findstr ^=') do set cpuModel=%%a
echo CPU%firstDelimiter%%cpuModel%%secondDelimiter% >> %outputFile%

REM GPU
for /f "delims=^= tokens=2" %%a in ('wmic path Win32_VideoController get Caption /format:list') do set gpuModel=%%a
echo GPU%firstDelimiter%%gpuModel%%secondDelimiter% >> %outputFile%

REM RAM
REM for /f %%a in ('powershell "Get-WmiObject Win32_PhysicalMemory | format-wide SMBIOSMemoryType"') do set ramType=%%a
for /f "delims=^= tokens=2" %%a in ('wmic memorychip get SMBIOSMemoryType /format:list') do set ramType=%%a
set "ramType=%ramType: =%"
if %ramType%==20 (
    set "ramType=DDR"
) else if %ramType%==21 (
    set "ramType=DDR2"
) else if %ramType%==24 (
    set "ramType=DDR3" 
) else if %ramType%==26 (
    set "ramType=DDR4"
) else (
    set "ramType=Unknown"
)
for /f "delims=^= tokens=2" %%a in ('wmic memorychip get speed /format:list ^| findstr ^=') do set ramSpeed=%%a
for /f %%a in ('powershell "(Get-WMIObject Win32_PhysicalMemory | Measure-Object Capacity -Sum).sum/1GB"') do set ramGigabyte=%%a
echo RAM%firstDelimiter%%ramType% %ramSpeed%MHz %ramGigabyte%GB%secondDelimiter% >> %outputFile%

REM disk
Setlocal Enabledelayedexpansion
for /f %%a in ('wmic diskdrive list brief ^| find /i "\\.\PHYSICALDRIVE" /c') do set diskSum=%%a

for /f "delims=^=c tokens=2" %%a in ('wmic diskdrive get Caption /format:list') do (set /a n=1 & if !n!==1 set "diskCaption1=%%a")
REM set "diskCaption1=%diskCaption1: =%"
for /f "delims=^=c tokens=2" %%a in ('wmic diskdrive get Caption /format:list') do (set /a n+=1 & if !n!==2 set "diskCaption2=%%a")
REM set "diskCaption2=%diskCaption2: =%"

for /f "delims=^=c tokens=2" %%a in ('wmic diskdrive get Size /format:list') do (set /a n=1 & if !n!==1 set "diskSizeByte1=%%a")
set "diskSizeByte1=%diskSizeByte1: =%"
set diskSizeGigabyte1=%diskSizeByte1:~0,-9%
for /f "delims=^=c tokens=2" %%a in ('wmic diskdrive get Size /format:list') do (set /a n+=1 & if !n!==2 set "diskSizeByte2=%%a")
set "diskSizeByte2=%diskSizeByte2: =%"
set diskSizeGigabyte2=%diskSizeByte2:~0,-9%

if %diskSum%==1 (
    echo Disk1%firstDelimiter%%diskCaption1% %diskSizeGigabyte1%GB%secondDelimiter% >> %outputFile%
) else ( 
    echo Disk1%firstDelimiter%%diskCaption1% %diskSizeGigabyte1%GB%secondDelimiter% >> %outputFile%
    echo Disk2%firstDelimiter%%diskCaption2% %diskSizeGigabyte2%GB%secondDelimiter% >> %outputFile%
    )
endlocal

REM Current User
Setlocal Enabledelayedexpansion
for /f "delims=>c tokens=1" %%a in ('quser') do (set /a n+=1 & if !n!==2 set "currentUsername=%%a")
set "currentUsername=%currentUsername: =%"
echo Username%firstDelimiter%%currentUsername%%secondDelimiter% >> %outputFile%
endlocal

echo Done!
echo=
pause
