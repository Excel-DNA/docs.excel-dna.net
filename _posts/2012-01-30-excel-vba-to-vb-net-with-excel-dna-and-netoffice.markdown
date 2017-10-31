---
layout: post
title: "Excel VBA to VB.NET with Excel-DNA and NetOffice"
date: 2012-01-30 13:29:00 -0000
permalink: /2012/01/30/excel-vba-to-vb-net-with-excel-dna-and-netoffice/
categories: samples
---
Excel-DNA is a great library to help ease the path from Excel VBA to VB.NET. Last year another part of the puzzle fell in place: I discovered [NetOffice][netoffice], a version-independent set of Office interop assemblies put together by Sebastian Lange. By referencing the NetOffice assemblies instead of the official Primary Interop Assemblies (PIA) for Office, an Excel-DNA add-in can target various Excel versions with a single add-in, and also ease distribution of the required interop assemblies, even packing them into the `.xll` add-in itself.

To explore how Excel-DNA and NetOffice can combine to convert a VBA add-in to VB.NET, I picked a small add-in made by [Robert del Vicario][robert-vicario] that does a risk analysis simulation inspired by the [Pallisade @RISK][palisade-risk] add-in. I took [Robert's original RiskGen VBA add-in][riskgen-vba], and created a new Excel-DNA add-in in VB.NET (I used Visual Studio, but the free [SharpDevelop][sharpdevelop] IDE should work fine too). I documented the steps along the way of creating the VB.NET project, making an add-in based on Excel-DNA and using NetOffice to help port the VBA code to VB.NET. The resulting document (RiskGen Port Log.docx) outlining exactly how I ported the add-in, with the new VB.NET-based [RiskGen.NET][riskgen-net] is also on Robert's site.

I'm also looking for some more examples of free/open source VBA add-ins to port to Excel-DNA. The best add-ins will contain a mix of user-defined functions and macros which use the Excel object model. Please post to the Google group or mail me directly if you have any suggestions.

And as always, if you need any support porting your Excel VBA add-ins to .NET using Excel-DNA, I'm happy to help on the [Excel-DNA Google group][excel-dna-group].

[netoffice]: https://github.com/netoffice
[robert-vicario]: http://rwdvc.posterous.com
[palisade-risk]: http://www.palisade.com/risk/
[riskgen-vba]: http://rwdvc.posterous.com/riskgen-test
[riskgen-net]: http://rwdvc.posterous.com/riskgen-in-vbnet
[sharpdevelop]: http://www.icsharpcode.net/OpenSource/SD/Download/#SharpDevelop4x
[excel-dna-group]: http://groups.google.com/group/exceldna
