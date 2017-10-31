---
layout: post
title: "Excel-DNA version 0.29 - RC available"
date: 2011-05-16 13:09:00 -0000
permalink: /2011/05/16/excel-dna-version-0-29-rc-available/
categories: uncategorized
---
I have posted a release candidate of Excel-DNA version 0.29 to the CodePlex site. The download is available at [http://exceldna.codeplex.com/releases/view/66405][release-0-29]. I will wait a week or two for some confirmation that this version works correctly before setting this release to "recommended" status and updating the NuGet package. Any results from your testing with this version would be very helpful.

Excel-DNA version 0.29 adds support for a number of specialized Excel features. The 64-bit version of Excel 2010 is fully supported, registration-free Custom Task Panes can be created under Excel 2007 and later, direct COM server integration can improve integration with legacy VBA code, and macros with parameters are registered. In addition, there are some features to improve the development and debugging workflow, and a few minor bugfixes. The complete change list is included below.

More information about the new features will be posted on the Excel-DNA website in the coming weeks. Any comments or questions are welcome on the Google group - [http://groups.google.com/group/exceldna][excel-dna-group] - or by contacting me directly.

As always, I greatly appreciate any feedback on this version, and on Excel-DNA in general.

-Govert

### Complete change list

* **BREAKING CHANGE!** Changed `SheetId` in the `ExcelReference` type to an `IntPtr` for 64-bit compatibility.
* Changed initialization - only create sandboxed `AppDomain` under .NET 4 (or if explicitly requested with `CreateSandboxedAppDomain="true"` attribute on `DnaLibrary` tag in .dna file).
* Fixed memory leak when getting `SheetId` for `ExcelReference` parameters.
* Fixed Ribbon `RunTagMacro` when no Workbook open.
* Added support for the 64-bit version of Excel 2010 with the .Net 4 runtime.
* Added Cluster-safe function support for Excel 2010 HPC Cluster Connector - mark functions as `IsClusterSafe=true`.
* Added `CustomTaskPane` support and sample.
* Added COM server support for RTD servers and other `ComVisible` classes. Mark `ExternalLibraries` and Projects as `ComServer="true"` in the .dna file. Supports `Regsvr32` registration or by calling `ComServer.DllRegisterServer`. Allows direct RTD and VBA object instantiation. Includes `TypeLib` registration and packing support.
* Added support for macros with parameters.
* Added `ArrayResizer` sample.
* Added C# 4 dynamic type sample.
* Added Path attribute to `SourceItem` tag to allow external source.
* Added `LoadFromBytes` attribute to `ExternalLibrary` tag to prevent locking of .dll.
* Added `/O` output path option to `ExcelDnaPack`.
* Added "before" option to CommandBars xml.
* Added `Int64` support for parameters and return values.

[release-0-29]: http://exceldna.codeplex.com/releases/view/66405
[excel-dna-group]: http://groups.google.com/group/exceldna
