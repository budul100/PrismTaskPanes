# PrismTaskPanes

Wrapper for PRISM based task panes in Microsoft Office

## DryIoc

The NuGet packages for DryIoc can be found on the following locations:

* Package for Microsoft Excel on [https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.Excel](https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.Excel)
* Package for Microsoft PowerPoint on [https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.PowerPoint](https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.PowerPoint)

These elements must be registered by using regasm.exe with the parameter /codebase.

## Common helpers

The relevant helper packages can be found on the following locations:

* Package for common values on [https://www.nuget.org/packages/budul.PrismTaskPanes.Commons](https://www.nuget.org/packages/budul.PrismTaskPanes.Commons)

## How-To

These two register functions must be added to each NetOffice addin:

```
[ComRegisterFunction]
public static void Register(Type type)
{
    RegisterFunction(type);

    ExcelProvider.RegisterProvider<AddIn>();
}

[ComUnregisterFunction]
public static void Unregister(Type type)
{
    UnregisterFunction(type);

    ExcelProvider.UnregisterProvider<AddIn>();
}
```