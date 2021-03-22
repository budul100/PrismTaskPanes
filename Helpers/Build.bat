SET CONFIGURATION=Debug

CALL Unregister.bat

PUSHD ..

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S bin') DO RMDIR /S /Q "%%G"
FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

POPD

dotnet build "..\PrismTaskPane.DryIoc\PrismTaskPane.DryIoc.Excel\PrismTaskPane.DryIoc.Excel.csproj"  --configuration %CONFIGURATION%
dotnet build "..\PrismTaskPanes.Regions\PrismTaskPanes.Regions.csproj"  --configuration %CONFIGURATION%

if %CONFIGURATION% == Debug (

	dotnet build "..\Examples\Excel\ExcelAddIn1\ExcelAddIn1.csproj"  --configuration %CONFIGURATION%
	dotnet build "..\Examples\Excel\ExcelAddIn2\ExcelAddIn2.csproj"  --configuration %CONFIGURATION%
)

CALL Register.bat

PAUSE

if %CONFIGURATION% == Debug (
	START excel.exe
)