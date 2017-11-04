---
layout: post
title: "Excel-DNA 0.32 Released"
date: 2014-05-04 15:46:00 -0000
permalink: /2014/05/04/excel-dna-0-32-released/
categories: release
---
I've posted a final release of Excel-DNA version 0.32 to CodePlex ([https://exceldna.codeplex.com/releases/view/119190][exceldna-v0-32]) and the NuGet package repository ([https://www.nuget.org/packages/Excel-DNA][exceldna-nuget]).

Excel-DNA 0.32 consolidates a large number of bug fixes and improvements that have accumulated over the last year. In particular, a number of edge cases that affect Excel-DNA add-ins under Excel 2013 have been addressed.

Native asynchronous functions, available under Excel 2010 and later, are now supported. Runtime registration of delegate functions and external retrieval of registration details will allow development of extension features without requiring changes to the Excel-DNA core runtime - see the ExcelDna.CustomRegistration project for examples of the dynamic registration: [https://github.com/Excel-DNA/CustomRegistration][custom-registration]

Excel-DNA 0.32 introduces one breaking change: integer parameter conversions are modified to be consistent with VBA. Fractional values passed to functions with integer parameters are converted using the round-to-even convention - as is the case for VBA functions. This issue is discussed in more detail at [http://excel-dna.net/2014/05/03/excel-dna-0-32-breaking-changes-to-integer-and-boolean-parameter-handling/][param-conversion]

See the [Distribution\ChangeLog.txt][exceldna-changelog] file for a complete list of changes in this version.

As always, I greatly appreciate any feedback on this version, and on Excel-DNA in general. Any comments or questions are welcome on the Google group or by contacting me directly.

To ensure future development of Excel-DNA, please make a donation via PayPal or arrange for a corporate support agreement. See [http://excel-dna.net/support/][exceldna-support] for details.

[exceldna-v0-32]: https://exceldna.codeplex.com/releases/view/119190
[exceldna-nuget]: https://www.nuget.org/packages/Excel-DNA
[custom-registration]: https://github.com/Excel-DNA/CustomRegistration
[param-conversion]: /2014/05/03/excel-dna-0-32-breaking-changes-to-integer-and-boolean-parameter-handling/
[exceldna-changelog]: https://exceldna.codeplex.com/SourceControl/latest#Distribution/ChangeLog.txt
[exceldna-support]: /support/
