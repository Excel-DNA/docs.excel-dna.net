---
layout: page
title: "External Links to Tutorials, Examples and Projects"
---

If you are a VBA developer interested in moving to .NET, you should start with [Patrick O'Beirne's detailed VBA to Excel-DNA migration guide](http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/).

## Various Samples and Tutorials
- As a comprehensive example using many of the Excel-DNA features, be inspired by the [Financial Analytics Suite (FinAnSu)](http://brymck.github.com/finansu/), an open-source C# add-in built by Bryan McKelvey.
- [Ross McLean](https://web.archive.org/web/20140902002824/http://www.blog.methodsinexcel.co.uk/2010/08/16/why-excel-dna/) has a series of posts on getting started with Excel-DNA.
- [Mikael Katajamäki shows how to use Microsoft Solver Foundation to build a curve fitting function for Excel](http://mikejuniperhill.blogspot.com/2013/06/using-ms-solver-foundation-and-c-in.html)
- [Mikael Katajamäki shows how to use C++/CLI code as a wrapper class for native (Quantlib based) C++ code and interfaced the C# client code to Excel by using Excel-DNA] http://mikejuniperhill.blogspot.com/2018/10/wilmott-software-interoperability-in.html
- [Simon Murphy - xlls with Excel-DNA](http://smurfonspreadsheets.wordpress.com/2010/02/18/xlls-with-exceldna/)
- [Ed Parcell - Numerical analysis in Excel using C# with Excel-DNA and AlgLib](https://web.archive.org/web/20100511213800/http://edparcell.posterous.com/tag/excel)
- [Mathias Brandewinder - Mutant Excel with .NET and Excel-DNA](http://www.clear-lines.com/blog/post/Mutant-Excel-and-Net-with-ExcelDNA.aspx)
- [Mathias Brandewinder - Supercharge Excel functions with Excel-DNA and .NET parallelism](http://www.clear-lines.com/blog/post/Supercharge-Excel-functions-with-ExcelDNA-and-Net-parallelism.aspx)
- [Mike Woodhouse - A third way: DNA?](http://grumpyop.wordpress.com/2009/11/25/a-third-way-dna/)
- [Patrick O'Beirne - From VBA to VB.NET using Excel-DNA](http://sysmod.wordpress.com/2012/02/06/from-vba-to-vb-net-using-exceldna/)
- [Doctor Torsten - Bring Excel 2010 to Speed: Remote UDFs with Excel 2010 and HPC Server 2008 R2](http://web.archive.org/web/20140831133544/http://doctortorsten.wordpress.com/2011/01/10/remoteudfs/)
- [Luca Bolognese - A trading/portfolio management Excel Add-in based on the books by Ralph Vince](https://www.lucabol.com/posts/2007-01-04-a-tradingportfolio-management-excel-add-in-based-on-the-books-by-ralph-vince/)
- [Supermab's series of blog posts introducing Excel-DNA to Japan (in Japanese)](http://supermab.com/wp/category/excel/)
- [Joao Morais - WCF client sample](http://blog.ilab8.com/2011/01/28/excel-dna/)
- [teramonagi - Using R from Excel using Excel-DNA](http://mockquant.blogspot.com/2011/07/another-way-to-use-r-in-excel-for-net.html) (Also check out the [F# R type provider.](https://bluemountaincapital.github.io/FSharpRProvider/))
- [Gert-Jan van der Kamp - Streaming real-time data to Excel](http://www.codeproject.com/Articles/662009/Streaming-realtime-data-to-Excel)

## Various F# Examples
- a [wrapper for the ExcelReference type and C API](https://gist.github.com/mndrake/5963107) with some cell access features,
- an [object handle wrapper](https://github.com/mndrake/ExcelObjectHandler), and
- a [WPF-based Custom Task Pane](https://github.com/mndrake/ExcelCustomTaskPane) for Excel.
- [Three samples, including a function using the R type provider](https://web.archive.org/web/20171228052128/http://luajalla.azurewebsites.net/excel-dna-three-stories/) by Natallie Baikevich.
- [Bram Jochems](https://web.archive.org/web/20140403050217/http://bramjochems.com/blog/2013/10/example-f-exceldna/) has published a wonderful [collection of finance-related functions on GitHub](https://github.com/bramjochems/MyExcelLib), as well as some details on [creating a ribbon menu with F#](https://web.archive.org/web/20160714194609/http://bramjochems.com/blog/2014/05/creating-ribbon-menu-exceldna-f/).
- Useful Range wrappers by Kit Eason: [Higher-Order Functions for Excel](http://www.fssnip.net/aV).

## External projects using Excel-DNA
- [Dodoni.net is a free/open-source library for quantitative finance and numerical computing.](https://dodoni.github.io/).
- [Cubicle Tools](https://cubicle.codeplex.com) is a collection of tools that extends Excel for analytical and rapid development purposes. It includes an object handler and an add-in distribution system.
- [EQ Finance - Analytics library for derivatives pricing and risk management](http://www.eqfltd.com/software)
- [Technoscience UK](http://excelxll.com/) has some interesting add-ins to mirror Excel data between PCs.
- [Niels Bosma -  SEOTools add-in (free, but not open source)](http://nielsbosma.se/projects/seotools/)
- [compute!box!](http://web.archive.org/web/20130616043202/http://computebox.wordpress.com/) allows real-time interchange of data between spreadsheets (via Azure Service Bus).
- This [Office icon gallery](https://archive.codeplex.com/?p=imagemso) has an Excel-based viewer.
- [Jon Nyman's FxToExcel add-in](https://github.com/jon49/FxToExcel) brings financial program data into Excel.
- [Stock Quote Add-In for Excel](https://github.com/jbaurle/PMStockQuote) provides access to the Yahoo financial data through an Excel-DNA add-in.
- [DB-Addin for Excel](https://rkapl123.github.io/DBAddin/) is an MS Excel Addin for retrieving Database data via userdefined functions into Excel and writing Data (DBMapper), executing generic DML (DBAction) and doing all this in Sequences (DBSequence).
- Alex Chirokov's **ACQ** add-in provides a library of interpolation routines for Excel. The add-in includes 1D and 2D interpolators, scatter plot smoothing and a Mersenne Twister random number generator. To have a closer look:
    - Find the current release on GitHub: [https://github.com/ratesquant/ACQ/releases][acq-releases]
    - With the main repository on GitHub at [https://github.com/ratesquant/ACQ][acq-repo]
    - A very clear introduction to the library, including some of it's advanced features, is posted on Code Project: [http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA][interpolation-article]
Features I like about the add-in (apart from it using Excel-DNA) include:
    - A liberal open-source license
    - A clear and authoritative implementation of a particular domain
    - Very nice example of using object handles - an interpolator is build from the data, and then used to interpolate many values. ACQ has a clean implementation and great example of this technique.
    - All the functions have a common prefix ("`=acq`..."), making them easy to find in the function list, and use with the Excel-DNA IntelliSense extension.
    - PS: ACQ has a bonus feature that implements a Sudoku solver (and generator)! See the write-up here: [Sudoku Solver in Excel using C# and Excel-DNA][sudoku-solver-article].
- I noticed a very nice add-in developed by [Bryan McKelvey][bryan-mckelvey] called [FinAnSu][finansu]. The whole add-in is generously available under the MIT open source license, and is a fantastic example of what can be built with Excel-DNA.
    - [FinAnSu][finansu] uses a ribbon interface to make the various functions and macros easy to find. The RTD server support is used to implement asynchronous data update functions, providing a live quote feed from Bloomberg, Google or Yahoo! And then there is a bunch of useful-looking financial functions. Here's a little preview:

        ![FinAnSu Quote Animated][finansu-quote-img]

    * Find the project on GitHub: [https://github.com/brymck/finansu][finansu], with detailed [documentation][finansu-docs].
    * You can browse through the [source code][finansu] online, and you can also download a copy of the whole project.

[acq-releases]: https://github.com/ratesquant/ACQ/releases
[acq-repo]: https://github.com/ratesquant/ACQ
[interpolation-article]: http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA
[sudoku-solver-article]: http://www.codeproject.com/Articles/1098156/Sudoku-Solver-in-Excel-using-Csharp-and-Excel-DNA
[bryan-mckelvey]: https://github.com/brymck
[finansu]: https://github.com/brymck/finansu
[finansu-quote-img]: /assets/finansu-quote-animated.gif "FinAnSu Quote Animated"
[finansu-docs]: http://brymck.github.com/finansu/
