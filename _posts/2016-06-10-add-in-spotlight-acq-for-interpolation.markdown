---
layout: post
title: "Add-in spotlight: ACQ for interpolation"
date: 2016-06-10 21:23:00 -0000
permalink: /2016/06/10/add-in-spotlight-acq-for-interpolation/
categories: uncategorized
---
This is the first in an occasional series of posts about interesting Excel-DNA based add-ins.

Alex Chirokov's **ACQ** add-in provides a library of interpolation routines for Excel. The add-in includes 1D and 2D interpolators, scatter plot smoothing and a Mersenne Twister random number generator.

To have a closer look:

* Find the current release on GitHub: [https://github.com/ratesquant/ACQ/releases][acq-releases]
* With the main repository on GitHub at [https://github.com/ratesquant/ACQ][acq-repo]
* A very clear introduction to the library, including some of it's advanced features, is posted on Code Project: [http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA][interpolation-article]

Features I like about the add-in (apart from it using Excel-DNA) include:

* A liberal open-source license
* A clear and authoritative implementation of a particular domain
* Very nice example of using object handles - an interpolator is build from the data, and then used to interpolate many values. ACQ has a clean implementation and great example of this technique.
* All the functions have a common prefix ("`=acq`..."), making them easy to find in the function list, and use with the Excel-DNA IntelliSense extension.

Thank you for publishing a great add-in, Alex.

PS: ACQ has a bonus feature that implements a Sudoku solver (and generator)! See the write-up here: [Sudoku Solver in Excel using C# and Excel-DNA][sudoku-solver-article].

[acq-releases]: https://github.com/ratesquant/ACQ/releases
[acq-repo]: https://github.com/ratesquant/ACQ
[interpolation-article]: http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA
[sudoku-solver-article]: http://www.codeproject.com/Articles/1098156/Sudoku-Solver-in-Excel-using-Csharp-and-Excel-DNA
