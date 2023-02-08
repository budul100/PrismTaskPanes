# PrismTaskPanes

Wrapper for PRISM based task panes in Microsoft Office

## DryIoc

The NuGet packages for DryIoc can be found on the following locations:

* Package for Microsoft Excel on [https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.Excel](https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.Excel)
* Package for Microsoft PowerPoint on [https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.PowerPoint](https://www.nuget.org/packages/budul.PrismTaskPanes.DryIoc.PowerPoint)

These elements must be registered by using regasm.exe with the parameter /codebase.

## View

The relevant view packages can be found on the following locations:

* Package for common values on [https://www.nuget.org/packages/budul.PrismTaskPanes.View](https://www.nuget.org/packages/budul.PrismTaskPanes.View)

## How-To

These two register functions must be added to each Excel addin:

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

These two register functions must be added to each PowerPoint addin:

```
[ComRegisterFunction]
public static void Register(Type type)
{
    RegisterFunction(type);

    PowerPointProvider.RegisterProvider<AddIn>();
}

[ComUnregisterFunction]
public static void Unregister(Type type)
{
    UnregisterFunction(type);

    PowerPointProvider.UnregisterProvider<AddIn>();
}
```