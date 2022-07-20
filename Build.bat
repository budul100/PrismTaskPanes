@echo off

SET AdditionalsDir=.\Additionals
SET HelperScripts=%AdditionalsDir%\Scripts
SET SetupScripts=%AdditionalsDir%\Setup
SET NuGetDir=.\_Out
SET ProjectPaths='.\DryIoc\DryIoc.Excel\DryIoc.Excel.csproj','.\DryIoc\DryIoc.PowerPoint\DryIoc.PowerPoint.csproj','.\Base\View\View.csproj'

echo.
echo ##### Create PrismTaskPanes #####
echo.

CHOICE /C dr /N /M "Shall the [d]ebug version or the [r]elease version be compiled?"
SET CONFIGSELECTION=%ERRORLEVEL%
echo.

if /i "%ERRORLEVEL%" == "2" GOTO RELEASE

:DEBUG

SET CONFIGURATION=Debug

echo.
echo Clean solution
echo.

CALL "%HelperScripts%\Unregister.bat"

GOTO BUILD

:RELEASE

SET CONFIGURATION=Release

CHOICE /C mb /N /M "Shall the [b]uild (x.x._X_.0) or the [m]inor version (x._X_.0.0) be increased?"
SET VERSIONSELECTION=%ERRORLEVEL%
echo.

if /i "%VERSIONSELECTION%" == "1" (
	echo.
	echo Update minor version
	echo.

	powershell "%SetupScripts%\Update_VersionMinor.ps1 -projectPaths %ProjectPaths%"
)


echo.
echo Kill Excel and PowerPoint
echo.

powershell -command "Stop-Process -Name "excel" -Force"
powershell -command "Stop-Process -Name "powerpnt" -Force"

echo.
echo Clean solution
echo.

CALL "%HelperScripts%\Unregister.bat"
CALL "%HelperScripts%\CleanFolders.bat"

GOTO BUILD

:BUILD

echo.
echo Build solution
echo.

dotnet build ".\PrismTaskPanes.sln" --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (

	echo.
	CALL "%HelperScripts%\Register.bat"
	
	echo.
	PAUSE

	START excel.exe
	START powerpnt.exe

) else (

	echo.
	echo Copy NuGet packages
	echo.
	
	mkdir %NuGetDir%
	del %NuGetDir%\*.nupkg
	
	for /R %cd% %%f in (*.nupkg) do copy %%f %NuGetDir%\

	echo.
	echo Update build version
	echo.

	powershell "%SetupScripts%\Update_VersionBuild.ps1 -projectPaths %ProjectPaths%"

	echo.
	PAUSE
)