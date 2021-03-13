---
layout: page
title: "Ribbon Customization"
---
## Setting Ribbon Properties

The Ribbon extensibility model is a bit unusual. There is no opportunity to set the 'label' or 'image' property of the button after it is created, but there are `getLabel` and `getImage` callbacks that you can set up.

To get Excel to refresh your control (or the whole Ribbon extension) you need to set an onLoad callback (on the customUI element) which receives an `IRibbonUI` interface for you to keep. This interface has two methods - `Invalidate` and `InvalidateControl` - which you call when a control should be refreshed.

Excel-DNA can help with the implementation of the getImage callback - call the `ExcelRibbon.LoadImage` method (probably as `base.LoadImage(imageId)` in your code) with the imageId of the picture you want to show - this way you can load the images you specify in the .dna file.

To create dynamically set up ribbons, you need to build the xml also in a callback method (GetCustomUI). The below example shows a class that handles dynamic ribbon elements (also handling menu visibility, dynamic screentip and image display, dropdowns and recursive menus) 
as well as context menus (cell, row, column etc.), utilization of an inbuilt commandbar.

```vb
Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Linq
Imports System.Xml.Linq

' Class MenuHandler handles all Menu related aspects (context menu for building/refreshing, "Load Config" tree menu for retrieving stored configuration files, etc.)</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits CustomUI.ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        Globals.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='AddinTab' label='Your Addin'>"
        customUIXml +=
        "<group id='AddinGroup' label='Addin settings'>" +
            "<dropDown id='envDropDown' label='Environment:' sizeString='1234567890123456' getEnabled='GetEnabledSelect' getSelectedItemIndex='GetSelectedEnvironment' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' getSupertip='getSelectedTooltip' onAction='selectEnvironment'/>" +
            "<buttonGroup id='buttonGroup1'>" +
                "<menu id='configMenu' label='Settings'>" +
                    "<button id='user' label='User settings' onAction='showAddinConfig' imageMso='ControlProperties' screentip='Show/edit user settings for DB Addin' />" +
                    "<button id='central' label='Central settings' onAction='showAddinConfig' imageMso='TablePropertiesDialog' screentip='Show/edit central settings for DB Addin' />" +
                    "<button id='addin' label='DBAddin settings' onAction='showAddinConfig' imageMso='ServerProperties' screentip='Show/edit standard Addin settings for DB Addin' />" +
                "</menu>" +
                "<button id='props' label='Workbook Properties' onAction='showCProps' getImage='getCPropsImage' screentip='Change custom properties relevant for DB Addin:' getSupertip='getToggleCPropsScreentip' />" +
                "<button id='designmode' label='Buttons' onAction='showToggleDesignMode' getImage='getToggleDesignImage' getScreentip='getToggleDesignScreentip'/>" +
            "</buttonGroup>" +
            "<dialogBoxLauncher><button id='dialog' label='About DBAddin' onAction='showAbout' screentip='Show Aboutbox with help, version information and project homepage'/></dialogBoxLauncher>" +
        "</group>"
        ' DBAddin Tools Group:
        customUIXml +=
        "<group id='DBAddinToolsGroup' label='DB Addin Tools'>" +
            "<buttonGroup id='buttonGroup2'>" +
                "<dynamicMenu id='DBConfigs' label='DB Configs' imageMso='QueryShowTable' screentip='DB Function Configuration Files quick access' getContent='getDBConfigMenu'/>" +
            "</buttonGroup>" +
        "</group>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml +=
        "<group id='DBModifGroup' label='Execute DBModifier'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                                "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                                "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "</group>"
        customUIXml += "</tab></tabs></ribbon>"
        ' Context menus for refresh, jump and creation: in cell, row, column and ListRange (area of ListObjects)
        customUIXml += "<contextMenus>" +
        "<contextMenu idMso ='ContextMenuCell'>" +
            "<button id='refreshDataC' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
             "<menu id='createMenu' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperC' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionC' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceC' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separator' />" +
                "<button id='DBListFetchC' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuPivotTable'>" +
            "<button id='refreshDataPT' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Copy'/>" +
            "<menuSeparator id='MySeparatorPT' insertBeforeMso='Copy'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuCellLayout'>" +
            "<button id='refreshDataCL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuCL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperCL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionCL' tag='DBAction' label='DBAction' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceCL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separatorCL' />" +
                "<button id='DBListFetchCL' tag='DBListFetch' label='DBListFetch' imageMso='GroupLists' onAction='clickCreateButton'/>" +
                "<button id='DBRowFetchCL' tag='DBRowFetch' label='DBRowFetch' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryPivotCL' tag='DBSetQueryPivot' label='DBSetQueryPivot' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryListObjectCL' tag='DBSetQueryListObject' label='DBSetQueryListObject' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorCL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuRow'>" +
            "<button id='refreshDataR' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuColumn'>" +
            "<button id='refreshDataZ' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuListRange'>" +
            "<button id='refreshDataL' label='refresh data (Ctl-Sh-R)' imageMso='Refresh' onAction='clickrefreshData' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperL' tag='DBMapper' label='DBMapper' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceL' tag='DBSeqnce' label='DBSequence' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "</contextMenus></customUI>"
        Return customUIXml
    End Function

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon
    ' display warning button icon on Cprops change if DBFskip is set
    Public Function getCPropsImage(control As CustomUI.IRibbonControl) As String
        If Globals.getCustPropertyBool("DBFskip", ExcelDnaUtil.Application.ActiveWorkbook) Then
            Return "DeclineTask"
        Else
            Return "AcceptTask"
        End If
    End Function

    ' display warning icon on log button if warning has been logged
    Public Function getLogsImage(control As CustomUI.IRibbonControl) As String
        If Globals.WarningIssued Then
            Return "IndexUpdate"
        Else
            Return "MailMergeStartLetters"
        End If
    End Function

    ' display state of designmode in screentip of dialogBox launcher
    ' returns screentip and the state of designmode
    Public Function getToggleCPropsScreentip(control As CustomUI.IRibbonControl) As String
        getToggleCPropsScreentip = ""
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Try
                Dim docproperty As Microsoft.Office.Core.DocumentProperty
                For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                    If Left$(docproperty.Name, 5) = "DBFC" Or docproperty.Name = "DBFskip" Or docproperty.Name = "doDBMOnSave" Or docproperty.Name = "DBFNoLegacyCheck" Then
                        getToggleCPropsScreentip += docproperty.Name + ":" + docproperty.Value.ToString + vbCrLf
                    End If
                Next
            Catch ex As Exception
                getToggleCPropsScreentip += "exception when collecting docproperties: " + ex.Message
            End Try
        End If
    End Function

    ' click on change props: show builtin properties dialog
    Public Sub showCProps(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogProperties).Show
            ' to check whether DBFskip has changed:
            Globals.theRibbon.InvalidateControl(control.Id)
        End If
    End Sub

    ' toggle designmode button
    Sub showToggleDesignMode(control As CustomUI.IRibbonControl)
        ' utilize an inbuilt commandbar to toggle design mode state
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            cbrs.ExecuteMso("DesignMode")
        Else
            Globals.ErrorMsg("Couldn't toggle designmode, because Designmode commandbar button is not available (no button?)", "DBAddin toggle Designmode", MsgBoxStyle.Exclamation)
        End If
        ' update state of designmode in screentip
        Globals.theRibbon.InvalidateControl(control.Id)
    End Sub

    ' display state of designmode in screentip of button
    ' returns screentip and the state of designmode
    Public Function getToggleDesignScreentip(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            Return "Designmode is currently " + IIf(cbrs.GetPressedMso("DesignMode"), "on !", "off !")
        Else
            Return "Designmode commandbar button not available (no button on sheet)"
        End If
    End Function

    ' display state of designmode in icon of button
    ' returns screentip and the state of designmode
    Public Function getToggleDesignImage(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            If cbrs.GetPressedMso("DesignMode") Then
                Return "ObjectsGroupMenuOutlook"
            Else
                Return "SelectMenuAccess"
            End If
        Else
            Return "SelectMenuAccess"
        End If
    End Function

    ' for environment dropdown to get the total number of the entries
    Public Function GetItemCount(control As CustomUI.IRibbonControl) As Integer
        Return Globals.environdefs.Length
    End Function

    ' for environment dropdown to get the label of the entries
    Public Function GetItemLabel(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ' for environment dropdown to get the ID of the entries
    Public Function GetItemID(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ' after selection of environment (using selectEnvironment) used to return the selected environment
    Public Function GetSelectedEnvironment(control As CustomUI.IRibbonControl) As Integer
        Return Globals.selectedEnvironment
    End Function

    ' tooltip for the environment select drop down
    Public Function getSelectedTooltip(control As CustomUI.IRibbonControl) As String
        If CBool(fetchSetting("DontChangeEnvironment", "False")) Then
            Return "DontChangeEnvironment is set, therefore changing the Environment is prevented !"
        Else
            Return "configured for Database Access in Addin config %appdata%\Microsoft\Addins\DBaddin.xll.config (or referenced central/user setting)"
        End If
    End Function

    ' whether to enable environment select drop down
    ' returns true if enabled
    Public Function GetEnabledSelect(control As CustomUI.IRibbonControl) As Integer
        Return Not CBool(fetchSetting("DontChangeEnvironment", "False"))
    End Function

    ' Choose environment (configured in registry with ConstConnString(N), ConfigStoreFolder(N))
    Public Sub selectEnvironment(control As CustomUI.IRibbonControl, id As String, index As Integer)
        Globals.selectedEnvironment = index
        Globals.initSettings()
        ' provide a chance to reconnect when switching environment...
        conn = Nothing
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Dim retval As MsgBoxResult = QuestionMsg("ConstConnString:" + Globals.ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder + vbCrLf + vbCrLf + "Refresh DBFunctions in active workbook to see effects?", MsgBoxStyle.YesNo, "Changed environment to: " + fetchSetting("ConfigName" + Globals.env(), ""))
            If retval = vbYes Then Globals.refreshDBFunctions(ExcelDnaUtil.Application.ActiveWorkbook)
        Else
            Globals.ErrorMsg("ConstConnString:" + Globals.ConstConnString + vbCrLf + "ConfigStoreFolder:" + ConfigFiles.ConfigStoreFolder, "Changed environment to: " + fetchSetting("ConfigName" + Globals.env(), ""), MsgBoxStyle.Information)
        End If
    End Sub

    ' show xll standard config (AppSetting), central config (referenced by App Settings file attr) or user config (referenced by CustomSettings configSource attr)
    Public Sub showAddinConfig(control As CustomUI.IRibbonControl)
        Dim theEditDBModifDefDlg As EditDBModifDef = New EditDBModifDef()
        theEditDBModifDefDlg.ShowDialog()
        Globals.theRibbon.Invalidate()
    End Sub

    ' dialogBoxLauncher of DBAddin settings group: activate about box
    Public Sub showAbout(control As CustomUI.IRibbonControl)
        Dim myAbout As AboutBox = New AboutBox
        myAbout.ShowDialog()
    End Sub

    ' on demand, refresh the DB Config tree
    Public Sub refreshDBConfigTree(control As CustomUI.IRibbonControl)
        Globals.initSettings()
        ConfigFiles.createConfigTreeMenu()
        Globals.ErrorMsg("refreshed DB Config Tree Menu", "DBAddin: refresh Config tree...", MsgBoxStyle.Information)
        Globals.theRibbon.Invalidate()
    End Sub

    ' get DB Config Menu from File
    Public Function getDBConfigMenu(control As CustomUI.IRibbonControl) As String
        If ConfigFiles.ConfigMenuXML = vbNullString Then ConfigFiles.createConfigTreeMenu()
        Return ConfigFiles.ConfigMenuXML
    End Function

    ' load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag)
    Public Sub getConfig(control As CustomUI.IRibbonControl)
        ConfigFiles.loadConfig(control.Tag)
    End Sub

    ' set the name of the DBModifType dropdown to the sheet name (for the WB dropdown this is the WB name)
    Public Function getDBModifTypeLabel(control As CustomUI.IRibbonControl) As String
        getDBModifTypeLabel = If(control.Id = "DBSeqnce", "DBSequence", control.Id)
    End Function

    ' create the buttons in the DBModif sheet dropdown menu
    ' returns the menu content xml
    Public Function getDBModifMenuContent(control As CustomUI.IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If Not Globals.DBModifDefColl.ContainsKey(control.Id) Then Return ""
            Dim DBModifTypeName As String = IIf(control.Id = "DBSeqnce", "DBSequence", IIf(control.Id = "DBMapper", "DB Mapper", IIf(control.Id = "DBAction", "DB Action", "undefined DBModifTypeName")))
            For Each nodeName As String In Globals.DBModifDefColl(control.Id).Keys
                Dim descName As String = IIf(nodeName = control.Id, "Unnamed " + DBModifTypeName, Replace(nodeName, DBModifTypeName, ""))
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                Dim superTipStr As String = IIf(control.Id = "DBSeqnce", "executes " + DBModifTypeName + " defined in docproperty: " + nodeName, IIf(control.Id = "DBMapper", "stores data defined in DBMapper (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), IIf(control.Id = "DBAction", "executes Action defined in DBAction (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName).getTargetRangeAddress(), "undefined superTip")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='" + imageMsoStr + "' onAction='DBModifClick' tag='" + control.Id + "' screentip='do " + DBModifTypeName + ": " + descName + "' supertip='" + superTipStr + "' />"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            Globals.ErrorMsg("Exception caught while building xml: " + ex.Message)
            Return ""
        End Try
    End Function

    ' show a screentip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)
    ' returns the screentip
    Public Function getDBModifScreentip(control As CustomUI.IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" + control.Id + ")"
    End Function

    ' to show the DBModif sheet button only if it was collected
    ' returns true if to be displayed
    Public Function getDBModifMenuVisible(control As CustomUI.IRibbonControl) As Boolean
        Try
            Return Globals.DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' DBModif button activated, do DB Mapper/DB Action/DB Sequence or define existing (CtrlKey pressed)
    Public Sub DBModifClick(control As CustomUI.IRibbonControl)
        ' reset noninteractive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = "" : hadError = False
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        If Not ExcelDnaUtil.Application.CommandBars.GetEnabledMso("FileNewDefault") Then
            Globals.ErrorMsg("Cannot execute DB Modifier while cell editing active !", "DB Modifier execution", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Try
            If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
                createDBModif(control.Tag, targetDefName:=nodeName)
            Else
                ' DB sequence actions (the sequence to be done) are stored directly in DBMapperDefColl, so different invocation here
                If Not (ExcelDnaUtil.Application.ActiveWorkbook.ReadOnlyRecommended And ExcelDnaUtil.Application.ActiveWorkbook.ReadOnly) Then
                    Globals.DBModifDefColl(control.Tag).Item(nodeName).doDBModif()
                Else
                    Globals.ErrorMsg("ReadOnlyRecommended is set on active workbook (being readonly), therefore all DB Modifiers are disabled !", "DB Modifier execution", MsgBoxStyle.Exclamation)
                End If
            End If
        Catch ex As Exception
            Globals.ErrorMsg("Exception: " + ex.Message + ",control.Tag:" + control.Tag + ",nodeName:" + nodeName, "DBModif Click")
        End Try
    End Sub

    ' context menu entries in Insert/Edit DBFunc/DBModif and Assign DBSheet: create DB function or DB Modification definition
    Public Sub clickCreateButton(control As CustomUI.IRibbonControl)
        ' check for existing DBMapper or DBAction definition and allow exit
        Dim activeCellDBModifName As String = DBModifs.getDBModifNameFromRange(ExcelDnaUtil.Application.ActiveCell)
        Dim activeCellDBModifType As String = Left(activeCellDBModifName, 8)
        If (activeCellDBModifType = "DBMapper" Or activeCellDBModifType = "DBAction") And activeCellDBModifType <> control.Tag And control.Tag <> "DBSeqnce" Then
            Globals.ErrorMsg("Active Cell already contains definition for a " + activeCellDBModifType + ", inserting " + IIf(control.Tag = "DBSetQueryPivot" Or control.Tag = "DBSetQueryListObject", "DBSetQuery", control.Tag) + " here will cause trouble !", "Inserting not allowed")
            Exit Sub
        End If
        If control.Tag = "DBListFetch" Then
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBListFetch("""","""",R[1]C,,,TRUE,TRUE,TRUE)"})
        ElseIf control.Tag = "DBRowFetch" Then
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBRowFetch("""","""",TRUE,R[1]C:R[1]C[10])"})
        ElseIf control.Tag = "DBSetQueryPivot" Then
            ' first create a dummy pivot table
            ConfigFiles.createPivotTable(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above listobject
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",R[1]C)"})
        ElseIf control.Tag = "DBSetQueryListObject" Then
            ' first create a dummy ListObject
            ConfigFiles.createListObject(ExcelDnaUtil.Application.ActiveCell)
            ' then create the DBSetQuery assigning the (yet to be filled) query to the above listobject
            ConfigFiles.createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, {"RC", "=DBSetQuery("""","""",RC[1])"})
        ElseIf control.Tag = "DBMapper" Or control.Tag = "DBAction" Or control.Tag = "DBSeqnce" Then
            If activeCellDBModifType = control.Tag Then  ' edit existing definition
                DBModifs.createDBModif(control.Tag, targetDefName:=activeCellDBModifName)
            Else                                         ' create new definition
                DBModifs.createDBModif(control.Tag)
            End If
        End If
    End Sub
#Enable Warning IDE0060

End Class


' procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu
Public Module ConfigFiles

    ' loads config from file given in theFileName, the actual creation (createFunctionsInCells) of the configs is left out here for simplicity
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer
        retval = QuestionMsg("Inserting contents configured in " + theFileName, MsgBoxStyle.OkCancel, "DBAddin: Inserting Configuration...", MsgBoxStyle.Information)
        If retval = vbCancel Then Exit Sub
        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then ExcelDnaUtil.Application.Workbooks.Add

        ' open file for reading
        Try
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName, Text.Encoding.Default)
            Do
                ItemLine = fileReader.ReadLine()
                ' ConfigArray: Configs are tab separated pairs of <RC location vbTab function formula> vbTab <...> vbTab...
                Dim ConfigArray As String() = Split(ItemLine, vbTab)
                ' insert the ConfigArray
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, ConfigArray)
            Loop Until fileReader.EndOfStream
            fileReader.Close()
        Catch ex As Exception
            Globals.ErrorMsg("Error (" + ex.Message + ") during filling items from config file '" + theFileName + "' in ConfigFiles.loadConfig")
        End Try
    End Sub

    ' the folder used to store predefined DB item definitions
    Public ConfigStoreFolder As String

    ' fixed max Depth for Ribbon
    Const maxMenuDepth As Integer = 5

    ' fixed max size for menu XML
    Const maxSizeRibbonMenu = 320000

    ' used to create menu and button ids
    Private menuID As Integer

    ' tree menu stored here
    Public ConfigMenuXML As String = vbNullString

    ' for correct display of menu
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"

    ' creates the Config tree menu by reading the menu elements from the config store folder files/subfolders
    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        If Not Directory.Exists(ConfigStoreFolder) Then
            Globals.ErrorMsg("No predefined config store folder '" + ConfigStoreFolder + "' found, please correct setting and refresh!")
            ConfigMenuXML = "<menu xmlns='" + xnspace.ToString() + "'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
        Else
            ' top level menu
            currentBar = New XElement(xnspace + "menu")
            ' add refresh button to top level
            button = New XElement(xnspace + "button")
            button.SetAttributeValue("id", "refreshConfig")
            button.SetAttributeValue("label", "refresh DBConfig Tree")
            button.SetAttributeValue("imageMso", "Refresh")
            button.SetAttributeValue("onAction", "refreshDBConfigTree")
            currentBar.Add(button)
            ' collect all config files recursively, creating submenus for the structure (see readAllFiles) and buttons for the final config files.
            menuID = 0
            readAllFiles(ConfigStoreFolder, currentBar)
            ExcelDnaUtil.Application.StatusBar = ""
            currentBar.SetAttributeValue("xmlns", xnspace)
            ' avoid exception in ribbon...
            ConfigMenuXML = currentBar.ToString()
            If ConfigMenuXML.Length > maxSizeRibbonMenu Then
                MsgBox("Too many entries in " + ConfigStoreFolder + ", can't display them in a ribbon menu ..")
                ConfigMenuXML = "<menu xmlns='" + xnspace.ToString() + "'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
            End If
        End If
    End Sub

    ' reads all files contained in rootPath and its subfolders (recursively) and adds them to the DBConfig menu (sub)structure (recursively). For folders contained in specialConfigStoreFolders, apply further structuring by splitting names on camelcase or specialConfigStoreSeparator
    ' param rootPath root folder to be searched for config files
    ' param currentBar current menu element, where submenus and buttons are added
    ' param Folderpath for sub menus path of current folder is passed (recursively)
    Private Sub readAllFiles(rootPath As String, ByRef currentBar As XElement, Optional Folderpath As String = vbNullString)
        Try
            Dim newBar As XElement = Nothing
            Static MenuFolderDepth As Integer = 1 ' needed to not exceed max. menu depth (currently 5)

            ' read all leaf node entries (files) and sort them by name to create action menus
            Dim di As DirectoryInfo = New DirectoryInfo(rootPath)
            Dim fileList() As FileSystemInfo = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
            If fileList.Length > 0 Then
                For i = 0 To UBound(fileList)
                    newBar = New XElement(xnspace + "button")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " + Left$(fileList(i).Name, Len(fileList(i).Name) - 4) + " in active cell")
                    newBar.SetAttributeValue("tag", rootPath + "\" + fileList(i).Name)
                    newBar.SetAttributeValue("label", Folderpath + Left$(fileList(i).Name, Len(fileList(i).Name) - 4))
                    newBar.SetAttributeValue("onAction", "getConfig")
                    currentBar.Add(newBar)
                Next
            End If

            ' read all folder xcl entries and sort them by name
            Dim DirList() As DirectoryInfo = di.GetDirectories().OrderBy(Function(fi) fi.Name).ToArray()
            If DirList.Length = 0 Then Exit Sub
            ' recursively build branched menu structure from dirEntries
            For i = 0 To UBound(DirList)
                ExcelDnaUtil.Application.StatusBar = "Filling DBConfigs Menu: " + rootPath + "\" + DirList(i).Name
                ' only add new menu element if below max. menu depth for ribbons
                If MenuFolderDepth < maxMenuDepth Then
                    newBar = New XElement(xnspace + "menu")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("label", DirList(i).Name)
                    currentBar.Add(newBar)
                    MenuFolderDepth += 1
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                    MenuFolderDepth -= 1
                Else
                    newBar = currentBar
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                End If
            Next
        Catch ex As Exception
            Globals.ErrorMsg("Error (" + ex.Message + ") in MenuHandler.readAllFiles")
        End Try
    End Sub

End Module

Public Module Globals

    ' warning state for display change
    Public WarningIssued as Boolean

    'currently selected environment for DB Functions, zero based (env -1) !!
    Public selectedEnvironment As Integer

    'reference object for the Addins ribbon
    Public theRibbon As CustomUI.IRibbonUI

    ' show Error message to User and log as warning (errors would pop up the trace information window)
    ' param LogMessage the message to be shown/logged
    ' param errTitle optionally pass a title for the msgbox here
    Public Sub ErrorMsg(LogMessage As String, Optional errTitle As String = "DBAddin Error", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Critical)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, EventLogEntryType.Warning, EventLogEntryType.Information), caller) ' to avoid popup of trace log in nonInteractive mode...
        If Not nonInteractive Then MsgBox(LogMessage, msgboxIcon + MsgBoxStyle.OkOnly, errTitle)
    End Sub

    ' logs Message of eEventType to System.Diagnostics.Trace
    ' param Message Message to be logged
    ' param eEventType event type: info, warning, error
    ' param caller reflection based caller information: module.method
    Private Sub WriteToLog(Message As String, eEventType As EventLogEntryType, caller As String)
        ' collect errors and warnings for returning messages in executeDBModif
        If eEventType = EventLogEntryType.Error Or eEventType = EventLogEntryType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf
        If nonInteractive Then
            Trace.TraceInformation("Noninteractive: {0}: {1}", caller, Message)
        Else
            Select Case eEventType
                Case EventLogEntryType.Information
                    Trace.TraceInformation("{0}: {1}", caller, Message)
                Case EventLogEntryType.Warning
                    Trace.TraceWarning("{0}: {1}", caller, Message)
                    WarningIssued = True
                    ' at Addin Start ribbon has not been loaded so avoid call to it here..
                    If Not theRibbon Is Nothing Then theRibbon.InvalidateControl("showLog")
                Case EventLogEntryType.Error
                    Trace.TraceError("{0}: {1}", caller, Message)
            End Select
        End If
    End Sub

End Module
```

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
