@echo off

SET BaseDir=%~dp0..\..

echo %BaseDir%
Pause

powershell ".\_PreBuild.ps1 -baseDir %BaseDir%"
powershell ".\_CleanFolders.ps1 -baseDir %BaseDir%"
