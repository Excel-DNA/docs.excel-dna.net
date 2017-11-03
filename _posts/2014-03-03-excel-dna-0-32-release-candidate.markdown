---
layout: post
title: "Excel-DNA 0.32 Release Candidate"
date: 2014-03-03 23:16:00 -0000
permalink: /2014/03/03/excel-dna-0-32-release-candidate/
categories: uncategorized, exceldna, nuget
---
I have posted to CodePlex ([https://exceldna.codeplex.com/releases/view/119190][release-v0-32-pre]) and the NuGet package manager ([https://www.nuget.org/packages/Excel-DNA/0.32.0-rc1][nuget-v0-32-rc1]) a release candidate of the next Excel-DNA version.

I hope to make a final release in the next few weeks, once I've had confirmation that this version works well on the various platforms and Excel versions.

Please test, and let me know of any problems or surprises you run into, or confirm what features, platforms and Excel versions work correctly.

The CodePlex download is structured as before, and for the NuGet package manager, you can upgrade to the pre-release version with:

{% highlight powershell %}
PM> Upgrade-Package Excel-DNA -Pre
{% endhighlight %}

---

**Excel-DNA 0.32** consolidates a large number of bug fixes and improvements that have accumulated over the last year. In particular, a number of edge cases that affect Excel-DNA add-ins under Excel 2013 have been addressed.

Native asynchronous functions, available under Excel 2010 and later, are now supported. Runtime registration of delegate functions and external retrieval of registration details will allow development of extension features without requiring changes to the Excel-DNA core runtime.

Excel-DNA 0.32 is compatible with version 0.30, and introduces no notable breaking changes. See the Distribution\ChangeLog.txt file for a complete change list.

As always, I greatly appreciate any feedback on this version, and on Excel-DNA in general. Any comments or questions are welcome on the [Google group][excel-dna-group] or by contacting me directly.

*To ensure future development of Excel-DNA, please make a donation via PayPal or arrange for a corporate support agreement. See [http://excel-dna.net/support/][excel-dna-support] for details*.

Thank you for your continued support,
Govert

[release-v0-32-pre]: https://exceldna.codeplex.com/releases/view/119190
[nuget-v0-32-rc1]: https://www.nuget.org/packages/Excel-DNA/0.32.0-rc1
[excel-dna-group]: http://groups.google.com/group/exceldna
[excel-dna-support]: /support/
