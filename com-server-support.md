---
layout: page
title: "COM Server Support"
---
Functions and macros created in an Excel-DNA add-in can be called directly from Excel VBA by using Application.Run(...). However, .NET also supports creating rich object models that are exported as COM libraries, which can be Tools->Referenced in VBA. Excel-DNA has some advanced support to host COM-exported objects from Excel-DNA add-ins, giving some advantages over the regular .NET ‘Register for COM interop’ hosting approach:

- COM objects that are created via the Excel-DNA COM server support will be active in the same AppDomain as the rest of the add-in, allowing direct shared access to static variables, internal caches etc.
- COM registration for classes hosted by Excel-DNA does not require administrative access (even when registered via RegSvr32.exe).
- Everything needed for the COM server can be packed in a single-file .xll add-in, including the type library used for IntelliSense support in VBA.
Mikael Katajamäki has written some detailed tutorial posts on his [Excel in Finance](http://mikejuniperhill.blogspot.com/) blog that explore this Excel-DNA feature, with detailed explanation, step-by-step instructions, screen shots and further links. See:

- [Interfacing C# and VBA with Excel-DNA (no intellisense support)](http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna-no.html)
- [Interfacing C# and VBA with Excel-DNA (with intellisense support)](http://mikejuniperhill.blogspot.com/2014/03/interfacing-c-and-vba-with-exceldna_16.html)

Note that these techniques would works equally well with code written in VB.NET, allowing you to port VB/VBA libraries to VB.NET with Excel-DNA and then use these from VBA.

----

COM visible classes in `ExternalLibrary` tags marked `ComServer="true"`, and COM visible classes that implement IRtdServer can be activated through the .xll directly. Even if the add-in is not loaded in Excel, such objects can be created in VBA.

These classes are (persistently) registered by calling `regsvr32 <MyAddin>.xll` or dynamically by the add-in (for example in an `AutoOpen` method) by calling `ComServer.DllRegisterServer()`, and
unregistered by `regsvr32 /u <MyAddin>.xll` or by `ComServer.DllUnregisterServer()`.

Following are short examples both in C-Sharp and VB, these only demonstrate the unreferenced (late-bound) technique:

```csharp
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using System.Runtime.InteropServices;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
[ProgId("ComAddin.FunctionLibrary")]
public class AccessibleFunctions
{
	public double add(double x, double y)
	{
		return x + y;
	}
}

[ComVisible(false)]
class ExcelAddin : IExcelAddIn
{
	public void AutoOpen()
	{
		ComServer.DllRegisterServer();
	}
	public void AutoClose()
	{
		ComServer.DllUnregisterServer();
	}
}
```

The same in VB.NET

```vb
Imports ExcelDna.Integration
Imports ExcelDna.ComInterop
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDispatch)>
<ProgId("ComAddin.FunctionLibrary")>
<ComVisible(True)>
Public Class AccessibleFunctions

    Public Function add(x As Double, y As Double)
        Return x + y
    End Function

End Class

<ComVisible(false)>
Public Class AddInEvents
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ComServer.DllUnregisterServer()
    End Sub
End Class
```

```xml
<DnaLibrary Name="ComAddin" RuntimeVersion="v4.0">
  <ExternalLibrary Path="ComAddin.dll" ComServer="true" />
</DnaLibrary>
```

Usage in VBA:

```vb
Option Explicit

Sub tester()
    Dim lib As Object: Set lib = CreateObject("ComAddin.FunctionLibrary")
    Debug.Print lib.Add(12, 13)
    Set lib = Nothing
End Sub
```

I've added a [ComAddinZip] package where the two examples are contained as Visual Studio 2019 solutions, after building (if the ExcelDNA NuGet packages are not automatically fetched, just enter the package manager, this should usually fetch them) start Excel by opening bin/release/ComAddin.xll or ComAddin64.xll (depending on your bitness) and open the Tester.xlsm.
There you find the above mentioned VBA procedure.

Such classes can be accessed directly as RTD servers or from VBA using `CreateObject("MyServer.ItsProgId")`, and will be loaded in the add-in's AppDomain.
(The add-in need not be loaded for registered classes to be accessed through COM.)

A type library (.tlb) can be created for the assembly using tlbexp.exe, and will be registered if available (if the `.tlb` is found next to the `.dll`). If the assembly is packed in the .xll, the type library will be packed too.

[ComAddinZip]: /assets/ComAddin.zip