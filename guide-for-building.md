---
layout: page
title: "Excel-DNA Building Guide"
---

To build Excel-DNA from source, create a single \<root> directory first.  

Build Requirement: Visual Studio 2022 Community/Professional with C++ and .NET Desktop development support, set the path to MSBuild.exe in \<root>\ExcelDna\MasterBuild\MasterBuild.cmd:  

`set MSBuildPath="C:\Program Files\Microsoft Visual Studio\2022\Community\Msbuild\Current\Bin\amd64\MSBuild.exe"`

You have to clone the following repositories into the \<root> directory:

- [ExcelDna](https://github.com/Excel-DNA/ExcelDna)
- [IntelliSense](https://github.com/Excel-DNA/IntelliSense)
- [Registration](https://github.com/Excel-DNA/Registration)
- [ExcelDnaDoc](https://github.com/Excel-DNA/ExcelDnaDoc)
- [DeveloperTools](tbd)

Then, just invoke the script MasterBuild.cmd in \<root>\ExcelDna\MasterBuild\MasterBuild.cmd


