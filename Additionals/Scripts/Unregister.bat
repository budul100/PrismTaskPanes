@ECHO off

SET ExamplesDir=%~dp0..\..\Examples\DryIoc
SET Target=net5.0-windows

regsvr32 /u /s "%CommonProgramFiles%\PrismTaskPanes\PrismTaskPanes.Host.comhost.dll"

regsvr32 /u /s "%ExamplesDir%\Excel\ExcelAddIn1\bin\Debug\%Target%\ExcelAddIn1.comhost.dll"
regsvr32 /u /s "%ExamplesDir%\Excel\ExcelAddIn2\bin\Debug\%Target%\ExcelAddIn2.comhost.dll"

regsvr32 /u /s "%ExamplesDir%\PowerPoint\PowerPointAddIn1\bin\Debug\%Target%\PowerPointAddIn1.comhost.dll"
regsvr32 /u /s "%ExamplesDir%\PowerPoint\PowerPointAddIn2\bin\Debug\%Target%\PowerPointAddIn2.comhost.dll"