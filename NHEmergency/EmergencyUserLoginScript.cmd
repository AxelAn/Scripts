rem @echo off
XCOPY \\wihan.nh\Freigabe\Emergency\PrepUser %USERPROFILE%\PrepUser /E /H /I /R /Y
%SYSTEMROOT%\System32\WindowsPowerShell\v1.0\Powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy ByPass -File %USERPROFILE%\PrepUser\Set-EmergencyUserEnvironment.ps1
