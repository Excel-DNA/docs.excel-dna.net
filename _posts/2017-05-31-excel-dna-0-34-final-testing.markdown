---
layout: post
title: "Excel-DNA version 0.34"
date: 2017-05-31 01:12:00 -0000
permalink: /2017/05/31/excel-dna-0-34-final-testing/
categories: uncategorized
---
Excel-DNA version 0.34 has been released and is now available.

* on NuGet (as package ExcelDna.Addin 0.34.6), and
* as a standalone download from GitHub ([https://github.com/Excel-DNA/ExcelDna/releases/tag/v0.34.6][release-v0-34-6]).

Please post any issues you run into to the Google group ([https://groups.google.com/forum/#!forum/exceldna][exceldna-google-group])

The easiest way to test the new version is by installing (or upgrading to) the `ExcelDna.AddIn` NuGet package.

---

Excel-DNA version 0.34 introduces a much improved build procedure for add-ins created using the NuGet package (thanks to a fantastic work by [@caioproiete][caioproiete]!)

This replaces the error-prone post-build steps we had with a custom build helper and allows easier build output customization (see [https://github.com/Excel-DNA/ExcelDna/wiki/Build-Output-Customization][build-customization]).

Various bug fixes and smaller improvements are also included in this version:
* Add `ExplicitExports="false"` to NuGet .dna file template
* Fix getting `Application` from `ProtectedViewWindow`
* Add attempts to get `Application` object from all windows of class `EXCEL7`.
* Fix `ExcelAsyncUtil.Observe` re-open restart - broken by other fixes in the previous version. Add option to not restart.
* Change `ExcelRtdServer.ConnectData` to be more careful about raising an update notice. Calls to `Topic.UpdateNotify` during the `ConnectData` overload are now always ignored. If the topic value is updated (through `Topic.UpdateValue`) during ConnectData, and the same value is returned from ConnectData, then no spurious `UpdateNotify` is raised. If the value returned from `ConnectData` differs from `Topic.Value`, `UpdateNotify` will still be raised.
* Allow `AccessViolation` exceptions to be caught under .NET 4.0 - change marshaling wrapper from `DynamicMethod` to `MethodBuilder`.
* Fix `QueueAsMacro` failure after paste live preview.
* Fix `AssemblyResolve` re-entrancy race condition.

---

To make a donation to the project, or to arrange for a corporate support agreement that lets you ensure Excel-DNA will live on, please visit the [Excel-DNA Support][exceldna-support] page.

Thank you for your continued support and enthusiasm towards the Excel-DNA project!

[release-v0-34-6]: https://github.com/Excel-DNA/ExcelDna/releases/tag/v0.34.6
[exceldna-google-group]: https://groups.google.com/forum/#!forum/exceldna
[caioproiete]: https://github.com/caioproiete
[build-customization]: https://github.com/Excel-DNA/ExcelDna/wiki/Build-Output-Customization
[exceldna-support]: /support/
