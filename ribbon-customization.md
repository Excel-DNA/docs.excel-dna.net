---
layout: page
title: "Ribbon Customization"
---
## Setting Ribbon Properties

The Ribbon extensibility model is a bit unusual. There is no opportunity to set the 'label' or 'image' property of the button after it is created, but there are `getLabel` and `getImage` callbacks that you can set up.

To get Excel to refresh your control (or the whole Ribbon extension) you need to set an onLoad callback (on the customUI element) which receives an `IRibbonUI` interface for you to keep. This interface has two methods - `Invalidate` and `InvalidateControl` - which you call when a control should be refreshed.

Excel-DNA can help with the implementation of the getImage callback - call the `ExcelRibbon.LoadImage` method (probably as `base.LoadImage(imageId)` in your code) with the imageId of the picture you want to show - this way you can load the images you specify in the .dna file.

To create dynamically set up ribbons, you need to build the xml also in a callback method (GetCustomUI). The [RibbonVB Example](https://github.com/Excel-DNA/Samples/tree/master/RibbonVB) in the Samples repository shows a project that handles dynamic ribbon elements (also handling menu visibility, dynamic screentip and image display, dropdowns and recursive menus) 
as well as context menus (cell, row, column etc.), utilization of an inbuilt commandbar and handling of VBE.Interop Command buttons.

Another VB Example is in the tutorials: [RibbonBasics](https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/RibbonBasics), a Csharp Example can be found in [Ribbon](https://github.com/Excel-DNA/Samples/tree/master/Ribbon).

If you look for the imageMso identifiers for the built-in ribbon controls, these can easily be looked up when customizing the ribbon (context menu "customize menu" in the ribbon) and hovering over the control of interest (the imageMso is shown in brackets after the control name).

## References

### Articles on MSDN:

- [Customizing the 2007 Office Fluent Ribbon for Developers (3 parts)](http://msdn.microsoft.com/en-us/library/aa338202(office.12).aspx)

- Making object wrappers to ease the scenario you have: [Custom Task Panes, the Office Fluent Ribbon, and Reusing VBA Code in the 2007 Office System](http://msdn.microsoft.com/en-us/library/bb194905(v=office.12).aspx)

#### Some Excel 2010 info:

- [Customizing the backstage view in Excel 2010](http://msdn.microsoft.com/en-us/library/ee815851(v=office.14).aspx)
- [Customizing context menus in Excel 2010](http://msdn.microsoft.com/en-us/library/ee691832(v=office.14).aspx)
- [Tab activation and scaling in Excel 2010](http://msdn.microsoft.com/en-us/library/ee691834(v=office.14).aspx)

### Other links:
- [Andy Pope's Ribbon Editor](http://www.andypope.info/vba/ribboneditor.htm) ([new, additional support for office 2010](http://www.andypope.info/vba/ribboneditor_2010.htm))
- [Ron de Bruin's site](https://www.rondebruin.nl/win/section2.htm) with details of the [Excel 2013 Backstage changes](http://www.rondebruin.nl/win/s2/win005.htm).
- [Discussion about Excel-DNA ribbons and the QAT](https://groups.google.com/forum/#!searchin/exceldna/qat/exceldna/hDocYCHy_Ao/SxnKUXDxiX8J).
- [Shareware tool ribboncreator](https://www.ribboncreator.de/en/)

Also note that a ribbon designed in VSTO can be exported to xml, which gives a `<customUI ...>` tag that can be used directly in Excel-DNA, though the ribbon handlers have to be re-implemented.
