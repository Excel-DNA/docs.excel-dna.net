---
layout: page
title: "CustomXMLParts"
---

After struggling with the limitations of custom docproperties (being generally much too short), I decided to focus on a technology that was introduced with Office 2007:
[CustomXmlParts](https://docs.microsoft.com/en-us/office/vba/api/office.customxmlparts)

CustomXmlParts allow you to store any arbitrary XML document in your Workbook (or Word document or Powerpoint presentation) being not visible to the end-user in the document itself.

In the following sample, I show some CustomXMLParts handling techniques from my [DB-Addin](https://github.com/rkapl123/DBAddin).

In my Addin, I use CustomXMLParts as an advanced storage for Database modifier definitions (for writing Excel Data to a Database table (DBMapper), doing DML Statements such as insert/update/delete (DBAction) and executing sequences of DBMappers and DBActions (DB Sequence)).

The [Sample](https://github.com/Excel-DNA/Samples/tree/master/CustomXMLParts) only contains creating and viewing/editing DBMapper Definitions in a stripped down class (all in one) without any other meaningful action code to concentrate on the usage of CustomXMLParts. Besides that, the usage of validating the XML against an existing schema is demonstrated as well.

## CustomXMLParts Usage

All CustomXMLParts Objects are found in the namespace `Microsoft.Office.Core`, referenced by the Office.dll within Exceldna.Interop.  

Fetching an existing CustomXMLParts XML document is done by selecting the required namespace:

```vbnet
        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
```

Adding the XML document if the namespace doesn't yet exist (in a new workbook) is done by adding the root element:
```VB
        If CustomXmlParts.Count = 0 Then
            ' in case no CustomXmlPart in Namespace DBModifDef exists in the workbook, add one
            ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        End If
```

### Adding sub elements

Sub elements are added with the Methond `AppendChildNode` of the selected node:

```vbnet
        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType, NamespaceURI:="DBModifDef")
```

The appended child element is placed last, to append further child elements, you need to call `LastChild`
```vbnet
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
        ' append the detailed settings to the definition element
        dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:=createdDBModifType + Guid.NewGuid().ToString())
        dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
```

### Retrieving elements

When retrieving element values, it's a good idea to check for the count of nodes contained to avoid exceptions:
```vbnet
        Dim nodeCount As Integer = definitionXML.SelectNodes("ns0:" + nodeName).Count
        If nodeCount = 0 Then
            getParamFromXML = "" ' optional nodes become empty strings
        Else
            getParamFromXML = definitionXML.SelectSingleNode("ns0:" + nodeName).Text
        End If
```

### Iterating through nodes

When iterating through nodes you take the `ChildNodes` method of the (root) node object and us `BaseName` of the iterator variable (node object) to get it's element name.
Here the name is usually in the (one and only) attribute "name" of the element, so if that exists, it is taken as the nodes name.

```vbnet
	For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
		Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
		If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
			Dim nodeName As String
			If customXMLNodeDef.Attributes.Count > 0 Then
				nodeName = customXMLNodeDef.Attributes(1).Text
			Else
				nodeName = customXMLNodeDef.BaseName + "unknown"
			End If

			' finally create the DBModif Object and fill parameters into CustomXMLPart:
			Dim newDBModif As DBModif = New DBModif(customXMLNodeDef, DBModiftype)
```
