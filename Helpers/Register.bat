@ECHO off

SET ExampleDirectoryExcel=..\Examples\Excel
SET ExampleDirectoryPowerPoint=..\Examples\PowerPoint

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryExcel%\ExampleAddIn1\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryExcel%\ExampleAddIn1\bin\Debug\net472\ExampleAddIn1.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryExcel%\ExampleAddIn2\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryExcel%\ExampleAddIn2\bin\Debug\net472\ExampleAddIn2.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryPowerPoint%\ExampleAddIn1\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryPowerPoint%\ExampleAddIn1\bin\Debug\net472\ExampleAddIn1.dll

"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryPowerPoint%\ExampleAddIn2\bin\Debug\net472\PrismTaskPanes.dll
"%windir%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" /tlb %ExampleDirectoryPowerPoint%\ExampleAddIn2\bin\Debug\net472\ExampleAddIn2.dll