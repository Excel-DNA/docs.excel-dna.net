---
layout: post
title: "Excel-DNA 0.33 Release Candidate and License Change"
date: 2015-08-03 00:50:00 -0000
permalink: /2015/08/03/excel-dna-0-33-release-candidate-and-license-change/
categories: uncategorized
---
## Version 0.33 Release Candidate

Excel-DNA 0.33 contains a number of bug fixes and improvements, including a [diagnostic logging][diagnostic-logging] approach based on the .NET Trace classes. More details can be found in the current [ChangeLog][changelog].

This version will also be the foundation for a first release of the custom Registration extension and the ongoing work towards on-sheet IntelliSense for user-defined functions.

A release candidate for the new version is available

* as a [GitHub release][github-release];
* as a [CodePlex download][codeplex-release];
* and in the (pre-release) [ExcelDna.AddIn NuGet package][addin-nupkg].

Please help me test that the new version works correctly in the many different ways, Excel and Windows versions, and languages where Excel-DNA add-ins run.

If you run into any unexpected behavior bugs or regressions, please post to the [Google group][excel-dna-group] or contact me directly.

Also, if you are able to confirm that the new version works in a particular setting, please post that too. Details about what functionality you've tested and what operating environment (including .NET and Excel version) you are running with, would help me a lot.


## License Change

~~For the Excel-DNA project, I've changed to the standard [MIT license][mit-license]. This has become the most common open-source license aligned with my intention of making Excel-DNA free for all use, including commercial use~~. **UPDATE 2015-09-03**: Excel-DNA is (again) licensed under the [zlib license][license]. More details [here][post-v0-33-8-rc2].

If you have any concerns with this change, please let me know.


## NuGet Packages

With this version, I am re-aligning the Excel-DNA package names on NuGet with the assembly names and standard naming conventions. The main packages for this release will be:

* **[ExcelDna.AddIn][addin-nupkg]** - Includes the `.xll` and creates a complete add-in when installed into a Class Library project. This is update of the "Excel-DNA" package.

* **[ExcelDna.Integration][integration-nupkg]** - Contains only the integration reference library, suitable for referencing in third-party libraries that are intended to be used in Excel-DNA add-ins. An update of the "Excel-DNA.Lib" package.

The old packages will be updated to refer to the new ones as dependencies, which should allow package updates to work correctly.


## GitHub

The Excel-DNA project is (slowly) moving to [GitHub][exceldna-github-home].

* The core library project can be found at [https://github.com/Excel-DNA/ExcelDna][exceldna-github-repo], where the latest source versions are hosted.
* The best documentation and links to related projects and other source is still found on the old [CodePlex site][exceldna-codeplex].
* For general questions and discussion about Excel-DNA, please continue use the [Excel-DNA Google group][excel-dna-group].
* Specific issues, bug reports and feature requests can be added to the [GitHub Issues][exceldna-github-issues] list.
* For a permanent bookmark to the project, please use the Excel-DNA home page at [http://excel-dna.net](/).

---

Thank you for your continued support of Excel-DNA!

[diagnostic-logging]: https://github.com/Excel-DNA/ExcelDna/wiki/Diagnostic-Logging
[changelog]: https://github.com/Excel-DNA/ExcelDna/blob/master/Distribution/ChangeLog.txt
[github-release]: https://github.com/Excel-DNA/ExcelDna/releases/tag/v0.33.7-rc1
[codeplex-release]: https://exceldna.codeplex.com/releases/view/616591
[addin-nupkg]: https://www.nuget.org/packages/ExcelDna.AddIn/
[excel-dna-group]: https://groups.google.com/forum/#!forum/exceldna
[license]: https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt
[post-v0-33-8-rc2]: /2015/09/03/excel-dna-version-0-33-8-rc2-available/
[mit-license]: https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt
[exceldna-github-home]: https://github.com/Excel-DNA
[exceldna-github-repo]: https://github.com/Excel-DNA/ExcelDna/
[exceldna-codeplex]: https://exceldna.codeplex.com
[exceldna-github-issues]: https://github.com/Excel-DNA/ExcelDna/issues/
[integration-nupkg]: https://www.nuget.org/packages/ExcelDna.Integration/
