@echo off

SET HelpersDir=.\Helpers
SET SetupScripts=%HelpersDir%\Scripts
SET SlnPaths='.\DryIoc\DryIoc.Excel\PrismTaskPanes.DryIoc.Excel.csproj','.\DryIoc\DryIoc.PowerPoint\PrismTaskPanes.DryIoc.PowerPoint.csproj','.\Base\PrismTaskPanes.Host\PrismTaskPanes.Host.csproj','.\Base\PrismTaskPanes.Commons\PrismTaskPanes.Commons.csproj'

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

CALL "%HelpersDir%\Unregister.bat"

GOTO BUILD

:RELEASE

SET CONFIGURATION=Release

CHOICE /C mb /N /M "Shall the [m]inor version (x._X_.0.0) or the [b]uild (x.x._X_.0) be increased?"
SET VERSIONSELECTION=%ERRORLEVEL%
echo.

if /i "%VERSIONSELECTION%" == "1" (
	echo.
	echo Update minor version
	echo.

	powershell "%SetupScripts%\Update_VersionMinor.ps1 -projectPaths %SlnPaths%"
)

echo.
echo Clean solution
echo.

CALL "%HelpersDir%\Unregister.bat"
CALL "%HelpersDir%\Clean.bat"

GOTO BUILD

:BUILD

echo.
echo Build solution
echo.

dotnet build ".\DryIoc\DryIoc.Excel\PrismTaskPanes.DryIoc.Excel.csproj" --configuration %CONFIGURATION%
dotnet build ".\DryIoc\DryIoc.PowerPoint\PrismTaskPanes.DryIoc.PowerPoint.csproj" --configuration %CONFIGURATION%

dotnet build ".\Base\PrismTaskPanes.Host\PrismTaskPanes.Host.csproj" --configuration %CONFIGURATION%
dotnet build ".\Base\PrismTaskPanes.Commons\PrismTaskPanes.Commons.csproj" --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (

	dotnet build ".\Examples\DryIoc\Excel\ExcelAddIn1\ExcelAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build ".\Examples\DryIoc\Excel\ExcelAddIn2\ExcelAddIn2.csproj" --configuration %CONFIGURATION%

	dotnet build ".\Examples\DryIoc\PowerPoint\PowerPointAddIn1\PowerPointAddIn1.csproj" --configuration %CONFIGURATION%
	dotnet build ".\Examples\DryIoc\PowerPoint\PowerPointAddIn2\PowerPointAddIn2.csproj" --configuration %CONFIGURATION%

	echo.
	CALL "%HelpersDir%\Register.bat"
	
	echo.
	PAUSE

	START excel.exe
	START powerpnt.exe

) else (

	del .\_NuGet\*.nupkg
	for /R %cd% %%f in (*.nupkg) do copy %%f .\_NuGet\

	if /i "%VERSIONSELECTION%" == "2" (
		echo.
		echo Update build version
		echo.

		powershell "%SetupScripts%\Update_VersionBuild.ps1 -projectPaths %SlnPaths%"
	)

	echo.
	PAUSE
)