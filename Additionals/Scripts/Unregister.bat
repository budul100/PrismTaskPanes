@ECHO off

SET ExampleDirectory=%~dp0..\..\Examples\DryIoc

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister /tlb "%ExampleDirectory%\Excel\ExcelAddIn1\bin\Debug\net472\ExcelAddIn1.dll"
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister /tlb "%ExampleDirectory%\Excel\ExcelAddIn2\bin\Debug\net472\ExcelAddIn2.dll"

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister /tlb "%ExampleDirectory%\PowerPoint\PowerPointAddIn1\bin\Debug\net472\PowerPointAddIn1.dll"
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /unregister /tlb "%ExampleDirectory%\PowerPoint\PowerPointAddIn2\bin\Debug\net472\PowerPointAddIn2.dll"
