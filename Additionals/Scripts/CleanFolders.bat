@echo off

SET BaseDir=%~dp0..\..

pushd %~dp0

REM powershell ".\_PreBuild.ps1 -baseDir %BaseDir%"
powershell ".\_CleanFolders.ps1 -baseDir %BaseDir%"

popd

pause