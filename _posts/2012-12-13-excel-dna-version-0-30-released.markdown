---
layout: post
title: "Excel-DNA Version 0.30 Released"
date: 2012-12-13 22:30:00 -0000
permalink: /2012/12/13/excel-dna-version-0-30-released/
categories: uncategorized, releases
---
Excel-DNA Version 0.30 has now been released on CodePlex. The download
is available from [http://exceldna.codeplex.com/releases/view/95861][release-v0-30].

This version implements support for RTD-based asynchronous worksheet
functions based on a thread-safe RTD base class - `ExcelRtdServer`. The
asynchronous support is designed to (optionally) integrate with
the .NET 4.0 Task-based operations, as well as the Reactive Extensions
library, allowing `IObservable`s to be exposed as "live" worksheet UDFs
- (thus "RxExcel"). The language-specific support for asynchronous
functions found in C# 5, Visual Basic 11 and F# 2.0 can be easily
integrated with the Excel-DNA asynchronous interfaces. Some examples
experimenting with the new async features are available in the
download.

Various bug fixes have also accumulated over the last 18 months, and
are collected in this release.

As always, I greatly appreciate any feedback on this version, and on
Excel-DNA in general. Any comments or questions are welcome on the
Google group or by contacting me directly.

## Changelist - Version 0.30 (12 December 2012)

* Fixed `LoadComAddIn` error when using a direct `ExcelComAddIn`-derived class.
* Fixed (Ribbon Helper) display in ribbon tooltips.
* Fixed RTD / array formula activation bug.
* Fixed `IsMacroType=true` reference argument sheet error (`ExcelReference` pointed to active sheet instead of current sheet).
* Fixed array marshaling pointer manipulation concern under 64-bit Excel.
* Fixed check for derived attributes too - for backward compatibility with v. 0.25.
* Fixed assembly multiple-loading problem for packed assemblies.
* Fixed persistent COM registration (`Regsv32.exe` / `ComServer.DllRegisterServer`) to allow `HKCR` registration whenever possible (for UAC elevation issue).
* Fixed Excel version check when COM / RTD Server loads before add-in is loaded - ribbon would not load.
* Fixed `IntPtr` `StackOverflowException` in high-memory 32-bit processes.
* Fixed custom task panes `UserControl` activation - do `HKCR` registration whenever possible (for UAC elevation issue).
* Fixed `double[0,1]` array marshaling memory allocation error with potential access violation.
* Allow abstract base classes in `ExcelRibbon` class hierarchy. Now loads the first concrete descendent of ExcelRibbon as the ribbon handler.
* Remove Obsolete class `ExcelDna.Integration.Excel`. (Use `ExcelDnaUtil` instead.) Allows smooth `XlCall` usage.
* Allow external `SourceItem` packing.
* Add `ExcelAsyncUtil` for async macro calls.
* Add thread-safe RTD server base class `ExcelRtdServer`.
* Add async function helper as `ExcelAsyncUtil.Run`.
* Add support for Reactive Extensions via RTD via `ExcelAsyncUtil.Observe` and related interfaces.
* Change `ExcelRibbon` and `ComAddIn` loading to use declared `ProgId` and `Guid` if _both_ attributes are present. Fixed Ribbon QAT issue.
* Revisit caching of `Application` object.
* Rename `ExcelDna.Integration.Integration` to `ExcelDna.Integration.ExcelIntegration`.
* Implement macro shortcuts (from `ExcelCommand` attributes).
* Changed re-open via File->Open to do full `AppDomain` unload and add-in reload.

[release-v0-30]: http://exceldna.codeplex.com/releases/view/95861
