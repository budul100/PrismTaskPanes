@ECHO off

SET ExamplesDir=%~dp0..\..\Examples\DryIoc
SET Target=net472

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "%ExamplesDir%\Excel\ExcelAddIn1\bin\Debug\%Target%\ExcelAddIn1.dll"
REM "%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "%ExamplesDir%\Excel\ExcelAddIn2\bin\Debug\%Target%\ExcelAddIn2.dll"

REM "%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "%ExamplesDir%\PowerPoint\PowerPointAddIn1\bin\Debug\%Target%\PowerPointAddIn1.dll"
REM "%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /codebase "%ExamplesDir%\PowerPoint\PowerPointAddIn2\bin\Debug\%Target%\PowerPointAddIn2.dll"
