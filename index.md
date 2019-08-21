---
layout: home
---
## Introduction
Excel-DNA is an independent project to integrate .NET into Excel. With Excel-DNA you can make native (.xll) add-ins for Excel using C#, Visual Basic.NET or F#, providing high-performance user-defined functions (UDFs), custom ribbon interfaces and more. Your entire add-in can be packed into a single .xll file requiring no installation or registration.


## Getting Started
If you are using a version of Visual Studio that supports the [NuGet Package Manager][nuget] (including Visual Studio 2012 Express for Windows Desktop or any more recent Visual Studio version), the easiest way to make an Excel-DNA add-in is to:

1. Create a new **Class Library (.NET Framework)** project in Visual Basic, C# or F#;

2. Use the **Manage NuGet Packages** dialog or the **Package Manager Console** to install the **Excel-DNA** package:
{% highlight powershell %}
PM> Install-Package ExcelDna.AddIn
{% endhighlight %}

{:start="3"}
3. Add your code (C#, Visual Basic.NET or F#):
{% highlight csharp %}
using ExcelDna.Integration;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string SayHello(string name)
    {
        return "Hello " + name;
    }
}
{% endhighlight %}

{:start="4"}
4. Compile, load and use your function in Excel:

{% highlight csharp %}
=SayHello("World!")
{% endhighlight %}

The [Excel-DNA NuGet Package][nuget-package] installs the required files and configures your project to build an Excel-DNA add-in.

Alternatively, get the full Excel-DNA [download from GitHub][releases], and work through the [Getting Started][getting-started] page. The download includes a step-by-step guide to making your first C# add-in, and more information is available on the [Documentation][documentation] page.


## More Details
Excel-DNA is developed using .NET, and users have to install the freely available .NET Framework runtime. The integration is by an Excel Add-In (.xll) that exposes .NET code to Excel. The user code can be in text-based (.dna) script files (C#, Visual Basic or F#), or compiled .NET libraries (.dll). Excel-DNA supports both the .NET runtime version 2.0 (which is used by .NET versions 2.0, 3.0 and 3.5) and version 4. Add-ins can target either version of the runtime, and concurrent loading of both runtime versions into an Excel instance is supported.

Excel versions '97 through 2010 can be targeted with a single add-in. Advanced Excel features are supported, including multi-threaded recalculation (Excel 2007 and later), registration-free RTD servers (Excel 2002 and later) and customized Ribbon interfaces (Excel 2007 and 2010). There is support for integrated Custom Task Panes (Excel 2007 and later), offloading UDF computations to a Windows HPC cluster (Excel 2010 and later), and for the 64-bit versions of Excel 2010 and 2013.

Most managed UDF assemblies developed for Excel Services can be exposed to the Excel client with no modification. (Please contact me if you are interested in this feature.)

Since Excel-DNA uses the Excel C API, porting C/C++ add-in code based on the Excel XLL SDK is very easy. (No more `XLOPER`s!)

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

## Latest Releases
The current version is [Excel-DNA 0.34][release-v0-34], released in May 2017, and includes various bug-fixes, some performance optimization, and a new implementation of the NuGet package build integration.

[nuget]: http://nuget.org
[nuget-package]: https://www.nuget.org/packages/ExcelDna.AddIn/
[releases]: https://github.com/Excel-DNA/ExcelDna/releases
[getting-started]: http://exceldna.codeplex.com/wikipage?title=Getting%20Started&referringTitle=Home
[documentation]: http://exceldna.codeplex.com/documentation?referringTitle=Home
[release-v0-34]: /2017/05/31/excel-dna-0-34-final-testing/
