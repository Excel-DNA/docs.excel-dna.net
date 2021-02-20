---
layout: page
title: "Overview"
---
- The core library project can be found on [GitHub](https://github.com/Excel-DNA/ExcelDna), where the latest source versions are hosted.
- For general questions and discussion about Excel-DNA, use the [Google group](https://groups.google.com/group/exceldna) or [Stack Overflow](http://stackoverflow.com/questions/tagged/excel-dna).
- Specific issues, bug reports and feature requests can be added to the [GitHub Issues](https://github.com/Excel-DNA/ExcelDna/issues) list.
- For more about Excel-DNA, see the introductory information below, and the searchable (back to February 2007) [Google group](https://groups.google.com/group/exceldna) history.
- For a permanent bookmark to the project, use the Excel-DNA home page at [http://excel-dna.net](http://excel-dna.net).

## Introduction

Excel-DNA is an independent project to integrate .NET into Excel. We hope it will be useful to Excel users who currently write VBA code for functions and macros, and would like to start using .NET. Also interested would be C/C++ based .xll add-in developers who want to use the .NET framework to develop their add-ins.

(For a bit more background about .NET and Excel-DNA, see [What and why? - An introduction to .NET and Excel-DNA](what-and-why-an-introduction-to-net-and-excel-dna.md)).

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

Excel-DNA is developed using .NET, and users have to install the freely available .NET Framework runtime. The integration is by an Excel Add-In (.xll) that exposes .NET code to Excel. The user code can be in text-based (.dna) script files (C#, Visual Basic or F#), or compiled .NET libraries (.dll). Excel-DNA supports both the .NET runtime version 2.0 (which is used by .NET versions 2.0, 3.0 and 3.5) and version 4. Add-ins can target either version of the runtime, and concurrent loading of both runtime versions into an Excel instance is supported.

Excel versions '97 through 2016 can be targeted with a single add-in. Advanced Excel features are supported, including multi-threaded recalculation (Excel 2007 and later), registration-free RTD servers (Excel 2002 and later) and customized Ribbon interfaces (Excel 2007 and 2010). There is support for integrated Custom Task Panes (Excel 2007 and later), offloading UDF computations to a Windows HPC cluster (Excel 2010 and later), and for the 64-bit versions of Excel 2010 and 2013.

Most managed UDF assemblies developed for Excel Services can be exposed to the Excel client with no modification. (Please contact us if you are interested in this feature)

The latest release - [Excel-DNA v1.1] - includes support for both RTD-based asynchronous worksheet functions (Excel 2002 and later) and native Excel asynchronous functions (Excel 2010 and later). The RTD-based asynchronous support is designed to (optionally) integrate with the .NET 4.0 Task-based operations, as well as the Reactive Extensions library, allowing IObservables to be exposed as 'live' worksheet UDFs - (thus 'RxExcel'). The language-specific support for asynchronous functions in C# 5, Visual Basic 11 and F# 2.0 can be easily integrated with the Excel-DNA asynchronous interfaces.

## Important Links
The home page for Excel-DNA is at [http://www.excel-dna.net](http://www.excel-dna.net).

The documentation is still sparse, but if you need any help, try the main Excel-DNA support forum on Google Groups, [https://groups.google.com/group/exceldna](https://groups.google.com/group/exceldna), where an extensive history of discussions can be found and searched through.

You are also welcome to contact us with questions, comments or suggestions.

## Getting Started

Get going with some first steps by following the [Getting Started](getting-started.md) page.

To make a C# add-in with Visual Studio consult the [step-by-step-csharp-add-in.docx](assets/step-by-step-csharp-add-in.docx) guide.

If you are using a version of Visual Studio that supports the [NuGet Package Manager](http://nuget.org) (including Visual Studio 2019 Community), the easiest way to make an Excel-DNA add-in is to:

- Create a new **Class Library (.NET Framework)** project in C#, F#, or Visual Basic.
- Use the **Manage NuGet Packages** dialog or the Package Manager Console to install the **[ExcelDna.AddIn](https://www.nuget.org/packages/ExcelDna.AddIn/)** package:

```
PM> Install-Package ExcelDna.AddIn
```

The [Excel-DNA NuGet Package](https://www.nuget.org/packages/ExcelDna.AddIn/) installs the required files and configures your project to build an Excel-DNA add-in.

Alternatively, get the full package [Excel-DNA Download](https://github.com/Excel-DNA/ExcelDna/releases) from GitHub, and work through the [Getting Started](getting-started.md) guide. The download includes a step-by-step guide to making your first C# add-in.

## [How-To Instructions](how-tos)

## Samples

Various sample projects and snippets related to Excel-DNA are available in the [Samples](https://github.com/Excel-DNA/Samples) repository.

Additional samples are available in the [Distribution/Samples](https://github.com/Excel-DNA/ExcelDna/tree/master/Distribution/Samples) folder and contains various .dna files, each of which is a self-contained add-in that exhibit some Excel-DNA functionality.

The .dna files are .xml files that can be edited with a regular text editor.

To run any of the sample .dna files, make a copy of the `Distribution\ExcelDna.xll` file, place it next to the .dna file, and rename to have the same prefix. E.g. to run `Optional.dna`, make a copy of `ExcelDna.xll` called `Optional.xll`, and double-click, or `File->Open` to load in Excel.

## Power Tools
- [ExcelDnaDoc](https://github.com/Excel-DNA/ExcelDnaDoc) provides tools to make help generation easier.
- [Registration](https://github.com/Excel-DNA/Registration) allows the automatic generation of parameter and function conversions, removing boiler-plate code for optional parameters, async functions etc.
- [IntelliSense](https://github.com/Excel-DNA/IntelliSense) add in-sheet IntelliSense for Excel UDFs.

## Community Projects

- [ExcelDna-Unpack](https://github.com/augustoproiete/exceldna-unpack) is a command-line utility to extract the contents of Excel-DNA add-ins packed with [ExcelDnaPack](exceldna-packing-tool.md)
- [ExcelDna.Abstractions](https://github.com/augustoproiete/exceldna-abstractions) facilitates mocking & unit testing of Excel-DNA Add-Ins
- [ExcelDna.Diagnostics.Serilog](https://github.com/augustoproiete/exceldna-diagnostics-serilog) integrates Excel-DNA Diagnostic Logging with your Serilog logging pipeling within your Excel-DNA Add-In
- [Serilog.Sinks.ExcelDnaLogDisplay](https://github.com/augustoproiete/serilog-sinks-exceldnalogdisplay) is a Serilog sink that writes events to Excel-DNA LogDisplay
- [Serilog.Enrichers.ExcelDna](https://github.com/augustoproiete/serilog-enrichers-exceldna) is a Serilog Enricher with properties from Excel-DNA Add-Ins

## [External Links to Tutorials, Examples and Projects](external-links)

## Support

There is a searchable record of more than 5000 messages on the [Excel-DNA Google Group](https://groups.google.com/group/exceldna).

There are many questions answered on Stack Overflow under the tag [`excel-dna`](http://stackoverflow.com/questions/tagged/excel-dna).

**Please don't hesitate to ask.** If you are stuck or need some help using Excel-DNA your questions really are very welcome - whether you are just getting started, or an Excel-DNA expert.

And if you could help put together some proper documentation, please contact me. We'd be happy to add you as an editor in this repository.

## Related Projects
- [NetOffice](https://github.com/netoffice) is a set of version-independent assemblies to allow Office integration through the COM automation interface. The NetOffice libraries can be used from an Excel-DNA add-in to ease version-independent Excel add-in development, and ease compatibility with VBA.
- Visual Studio Tools for Office (VSTO) is Microsoft's preferred plan for integrating .NET with Office. It is mainly aimed at making it easy for Visual Studio developers to create solutions integrated with the Office applications. In contrast, Excel-DNA is (eventually) aimed at Excel end-users, as a compelling replacement for VBA, completely independent of Visual Studio.
- [Add-in Express](http://www.add-in-express.com) is a commercial alternative to VSTO for users with Visual Studio. It support making add-ins for the various Office products, not just Excel, and has helpful wizards and graphics designers.
- Jens Thiel's ManagedXll was an established, commercial product to easily create .xll libraries in .NET. If ManagedXll were free, Excel-DNA would not exist.
- [Statfactory's NeXL](https://statfactory.wordpress.com) are F# based connectors to get data from various platforms (Bloomberg, Quandl, Worldbank, IMF and the R language) into Excel.
- For making Excel Add-Ins in Python, have a look at [PyXLL](http://www.pyxll.com).
- There are a number of C/C++ libraries and tools that make creating .xlls easier than using the [Excel SDK](https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit) directly:
    * I initially used the [XLW](http://xlw.sourceforge.net/) open-source library.
    * The [XLL+ toolkit](https://www.planatechsolutions.com/xllplus/) is a highly regarded commercial option.
    * Keith Lewis has some modern C++ libraries for making .xlls, available on CodePlex at [https://archive.codeplex.com/?p=xll](https://archive.codeplex.com/?p=xll).

## Performance
Information about the performance of Excel-DNA user-defined functions can be found on the [Excel-DNA Performance](exceldna-performance.md) page.

## Formal Support Agreements
Corporate users of Excel-DNA, using the library as part of their mission critical infrastructure, are encouraged to enter into a formal support arrangement. We offer an annual subscription-based technical support agreement, providing direct support, priority bug-fixes and feature development and ensuring that Excel-DNA will continue to be updated and developed.

## Donations
Financial support for the Excel-DNA project encourages future development and is greatly appreciated. Transactions are processed by PayPal.
[![Donate via PayPal][paypal-image]][paypal-link]

## More Details
Excel-DNA is developed using .NET, and users have to install the freely available .NET Framework runtime. The integration is by an Excel Add-In (.xll) that exposes .NET code to Excel. The user code can be in text-based (.dna) script files (C#, Visual Basic or F#), or compiled .NET libraries (.dll). Excel-DNA supports both the .NET runtime version 2.0 (which is used by .NET versions 2.0, 3.0 and 3.5) and version 4. Add-ins can target either version of the runtime, and concurrent loading of both runtime versions into an Excel instance is supported.

Excel versions '97 through 2016 can be targeted with a single add-in. Advanced Excel features are supported, including multi-threaded recalculation (Excel 2007 and later), registration-free RTD servers (Excel 2002 and later) and customized Ribbon interfaces (Excel 2007 and 2010). There is support for integrated Custom Task Panes (Excel 2007 and later), offloading UDF computations to a Windows HPC cluster (Excel 2010 and later), and for the 64-bit versions of Excel 2010 and 2013.

Most managed UDF assemblies developed for Excel Services can be exposed to the Excel client with no modification. (Please contact me if you are interested in this feature.)

Since Excel-DNA uses the Excel C API, porting C/C++ add-in code based on the Excel XLL SDK is very easy. (No more `XLOPER`s!)

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

Originally, the project was hosted on [https://exceldna.codeplex.com](https://exceldna.codeplex.com), where you can still download the site in it's historic state as a package. After CodePlex' shutdown the archive site is however mostly unusable by now.

## Latest Releases
The current version is [Excel-DNA v1.1], released in June 2020 and includes numerous improvements and bug-fixes.

[Excel-DNA v1.1]: https://excel-dna.net/2020/06/29/excel-dna-version-1-1/
[paypal-link]: https://www.paypal.com/cgi-bin/webscr?cmd=_donations&amp;business=92N99RV5NQ29C&amp;lc=US&amp;item_name=Govert%20van%20Drimmelen&amp;item_number=ExcelDna&amp;currency_code=USD&amp;bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHosted
[paypal-image]: https://www.paypal.com/en_GB/i/btn/btn_donateCC_LG.gif "Donate via PayPal"
