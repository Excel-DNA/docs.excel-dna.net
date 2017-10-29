---
layout: post
title: "Financial Analytics Suite (FinAnSu) - made with Excel-DNA"
date: 2011-04-28 22:51:00 -0000
permalink: /2011/04/28/financial-analytics-suite-finansu-made-with-excel-dna/
categories: uncategorized
---
I recently noticed a very nice add-in developed by [Bryan McKelvey][bryan-mckelvey] called [FinAnSu][finansu]. The whole add-in is generously available under the MIT open source license, and is a fantastic example of what can be built with Excel-DNA.

[FinAnSu][finansu] uses a ribbon interface to make the various functions and macros easy to find. The RTD server support is used to implement asynchronous data update functions, providing a live quote feed from Bloomberg, Google or Yahoo! And then there is a bunch of useful-looking financial functions. Here's a little preview:

![FinAnSu Quote Animated][finansu-quote-img]

Find the project on Google code: [http://code.google.com/p/finansu/][finansu], with detailed documentation on the wiki: [http://code.google.com/p/finansu/wiki/Introduction][finansu-docs].

You can browse through the [source code][finansu-source] online, but to download a copy of the whole project you'll need a [Mercurial client][mercurial-client]. I just installed the one called Mercurial 1.8.2 MSI installer and ran `hg clone https://finansu.googlecode.com/hg/ finansu` from a command prompt.

[bryan-mckelvey]: http://www.brymck.com/
[finansu]: http://code.google.com/p/finansu/
[finansu-quote-img]: /images/finansu-quote-animated.gif "FinAnSu Quote Animated"
[finansu-docs]: http://code.google.com/p/finansu/wiki/Introduction
[finansu-source]: http://code.google.com/p/finansu/source/browse/
[mercurial-client]: http://mercurial.selenic.com/downloads/
