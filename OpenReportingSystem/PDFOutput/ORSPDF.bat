REM echo off

REM %1 = sourceFileName
REM %2 = targetFileName

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Command "&{\OpenReportingSystem\PDFOutput\ORSPDF.ps1 %1 %2}"