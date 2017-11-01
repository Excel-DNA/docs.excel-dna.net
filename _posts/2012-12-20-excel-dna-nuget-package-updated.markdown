---
layout: post
title: "Excel-DNA NuGet Package Updated"
date: 2012-12-20 22:30:00 -0000
permalink: /2012/12/20/excel-dna-nuget-package-updated/
categories: uncategorized, exceldna, nuget
---
I've updated and improved the [Excel-DNA NuGet package][nuget-package]. ([NuGet][nuget] is the Visual Studio package manager that makes it easy to download and install external libraries into your projects.)

To turn your Class Library project into an Excel add-in, just open Tools -> Library Package Manager -> Package Manager Console, and enter

{% highlight powershell %}
PM> Install-Package Excel-DNA
{% endhighlight %}

The Excel-DNA package now has an install script that creates the required `.dna` file, and a post-build step to copy the `.xll` and run the packing tool, and even configures debugging. The package should work for C#, Visual Basic and F# Class Library projects on:

* Visual Studio 2010 Professional and higher
* Visual Studio 2012 Professional and higher
* Visual Studio 2012 Express for Windows Desktop

Please post any feedback - bugs, good or bad comments - to the [Excel-DNA Google group][exceldna-group].

[nuget-package]: http://nuget.org/packages/Excel-DNA/
[nuget]: http://nuget.org
[exceldna-group]: http://groups.google.com/group/exceldna
