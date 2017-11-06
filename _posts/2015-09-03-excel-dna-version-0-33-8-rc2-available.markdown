---
layout: post
title: "Excel-DNA version 0.33.8-rc2 available"
date: 2015-09-03 17:25:39 -0000
permalink: /2015/09/03/excel-dna-version-0-33-8-rc2-available/
categories: uncategorized
---
I've made an updated release candidate available - this is version 0.33.8-RC2.

In this update:

* Excel-DNA is (again) licensed under the [zlib license][license] (thanks to [Chel][post-license-issue] for raising the potential ambiguity in the MIT license).

* There are improvements in the NuGet install scripts (thanks to [Caio Proiete][caioproiete] for the encouragement and help with this).

* There are some minor fixes, including the `XlCall` fix for the error: "_Cannot Inherit from sealed class XlCall_"

* The VC++ native project was updated to use the VS 2013 tools (allowing the project to build with the AppVeyor continuous integration service).

The update can be installed by:

* Updating the Excel-DNA NuGet package to pre-release version 0.33.8-rc2

* Installing the new (pre-release) `ExcelDna.AddIn` package version 0.33.8-rc2

* Downloading the binary (and source) distribution from CodePlex ([https://exceldna.codeplex.com/releases/view/616591][release-v0-33-8-rc2]).

**Please help me test this update!**

**You can notify me of any problems or questions you encounter, either through the [Google group][exceldna-google-group] or directly via email to <govert@icon.co.za>**.

**Please also confirm if it works!**
**What Windows and Excel version, and what features do you use?**

---

Thank you for your continued support of Excel-DNA.

[license]: https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt
[post-license-issue]: https://groups.google.com/forum/#!topic/exceldna/CRsJrQ6mJTo
[caioproiete]: https://github.com/caioproiete
[release-v0-33-8-rc2]: https://exceldna.codeplex.com/releases/view/616591
[exceldna-google-group]: https://groups.google.com/forum/#!forum/exceldna
