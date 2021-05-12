---
layout: page
title: "Checking for updates and download"
---

Following is a simple method to check for available updates of Add-ins/Programs:

The procedure can be called on start up of Excel or on displaying an About Dialog-box (as I decided in my case being the less intrusive variant). The parameter `doUpdate` decides whether only the check for the new version is performed or the new version is actually downloaded.
The update check/download requires a continuously increasing version number being available via a URL (here on githubs tag/release archive) of the form `https://domain.name/path/1.0.0.<release>`.
A local update folder can also be provided to allow for a central update by an administrator.

The UserMsg and QuestionMsg are just wrappers of `MsgBox` providing a headless mode (without pop-ups) and logging so you can replace them with your choice of `MsgBox(theMessage, msgboxIcon + questionType, questionTitle)`.

First, the necessary settings are fetched, you can also hard-code the default values (being the second argument of `fetchSetting`, described in more detail in [User settings and the .xll.config file](user-settings-and-the-xllconfig-file.md))
```vbnet
    Public Sub checkForUpdate(doUpdate As Boolean)
        Const updateFilenameZip = "downloadedVersion.zip"
        Dim localUpdateFolder As String = fetchSetting("localUpdateFolder", "")
        Dim localUpdateMessage As String = fetchSetting("localUpdateMessage", "A new version is available in the local update folder, after quitting Excel (is done next) start deployAddin.cmd to install it.")
        Dim updatesMajorVersion As String = fetchSetting("updatesMajorVersion", "1.0.0.")
        Dim updatesDownloadFolder As String = fetchSetting("updatesDownloadFolder", "C:\temp\")

        ' put your UrlBase here, where the release zip files can be found
        Dim updatesUrlBase As String = fetchSetting("updatesUrlBase", "https://github.com/rkapl123/DBAddin/archive/refs/tags/")
        Dim response As Net.HttpWebResponse = Nothing
        Dim urlFile As String = ""
```

Then, the procedure checks check for the zip file of the next higher revisions until no higher version can be found. `Net.SecurityProtocolType.Tls12` is only available starting with .NET 4.5
```vbnet
        Dim curRevision As Integer = My.Application.Info.Version.Revision
        ' try with highest possible Security protocol
        Try
            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 Or Net.SecurityProtocolType.SystemDefault
        Catch ex As Exception
            UserMsg("Error setting the SecurityProtocol: " + ex.Message())
            Exit Sub
        End Try
        ' always accept url certificate as valid
        Net.ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidationCallbackHandler

        Do
            urlFile = updatesUrlBase + updatesMajorVersion + (curRevision + 1).ToString() + ".zip"
            Dim request As Net.HttpWebRequest
            Try
                request = Net.WebRequest.Create(urlFile)
                response = Nothing
                request.Method = "HEAD"
                response = request.GetResponse()
            Catch ex As Exception
            End Try
            If response IsNot Nothing Then
                curRevision += 1
                response.Close()
            End If
        Loop Until response Is Nothing
```

get out, if no newer version could be found. In my case, I set a TextBox (`TextBoxDescription`) and a button (`CheckForUpdates`) to notify the user. for the notification step (`doUpdate = False`) stop here and let the user decide by clicking on the button `CheckForUpdates`.
```vbnet
        If curRevision = My.Application.Info.Version.Revision Then
            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "You have the latest version (" + updatesMajorVersion + curRevision.ToString() + ")."
            Me.TextBoxDescription.BackColor = Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)
            Me.CheckForUpdates.Text = "no Update ..."
            Me.CheckForUpdates.Enabled = False
            Me.Refresh()
            Exit Sub
        Else
            Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "A new version (" + updatesMajorVersion + curRevision.ToString() + ") is available " + IIf(localUpdateFolder <> "", "in " + localUpdateFolder, "on github")
            Me.TextBoxDescription.BackColor = Drawing.Color.DarkOrange
            Me.CheckForUpdates.Text = "get Update ..."
            Me.CheckForUpdates.Enabled = True
            Me.Refresh()
            If Not doUpdate Then Exit Sub
        End If
```

If there is a maintained local update folder, open it and let user update from there.
```vbnet
        If localUpdateFolder <> "" Then
            Try
                If QuestionMsg(localUpdateMessage, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    System.Diagnostics.Process.Start("explorer.exe", localUpdateFolder)
                    Me.quitExcelAfterwards = True
                    Me.Close()
                End If
            Catch ex As Exception
                UserMsg("Error when opening local update folder: " + ex.Message())
            End Try
            Exit Sub
        End If
```

Otherwise continue and download newest version. Progress information is put into the TextBoxDescription
```vbnet
        urlFile = updatesUrlBase + updatesMajorVersion + curRevision.ToString() + ".zip"

        ' create the download folder
        Try
            IO.Directory.CreateDirectory(updatesDownloadFolder)
        Catch ex As Exception
            UserMsg("Couldn't create file download folder (" + updatesDownloadFolder + "): " + ex.Message())
            Exit Sub
        End Try

        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Downloading new version from " + urlFile
        Me.Refresh()
        ' get the new version zip-file
        Dim requestGet As Net.HttpWebRequest = Net.WebRequest.Create(urlFile)
        requestGet.Method = "GET"
        Try
            response = requestGet.GetResponse()
        Catch ex As Exception
            UserMsg("Error when downloading new version: " + ex.Message())
            Exit Sub
        End Try
        ' save the version as zip file
        If response IsNot Nothing Then
            Dim receiveStream As Stream = response.GetResponseStream()
            Using downloadFile As IO.FileStream = File.Create(updatesDownloadFolder + updateFilenameZip)
                receiveStream.CopyTo(downloadFile)
            End Using
        End If
        response.Close()
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "Extracting " + urlFile + " to " + updatesDownloadFolder
        Me.Refresh()
        ' now extract the downloaded file and open the Distribution folder, first remove any existing folder...
        Try
            Directory.Delete(updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString(), True)
        Catch ex As Exception : End Try
        Try
            IO.Compression.ZipFile.ExtractToDirectory(updatesDownloadFolder + updateFilenameZip, updatesDownloadFolder)
        Catch ex As Exception
            UserMsg("Error when extracting new version: " + ex.Message())
        End Try
        Me.TextBoxDescription.Text = My.Application.Info.Description + vbCrLf + vbCrLf + "New version in " + updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString() + "\Distribution, start deployAddin.cmd to install the new Version."
        Me.Refresh()
```

Finally, open the windows explorer to let the user start the update process. After leaving the procedure, a hint to close Excel (or do this automatically) is probably a good idea (the Addin won't be deployed as long Excel is open)...
```vbnet
        Try
            System.Diagnostics.Process.Start("explorer.exe", updatesDownloadFolder + "DBAddin-" + updatesMajorVersion + curRevision.ToString() + "\Distribution")
        Catch ex As Exception
            UserMsg("Error when opening Distribution folder of new version: " + ex.Message())
        End Try
    End Sub
```

The ValidationCallbackHandler just returns true, if some more checks are needed put them here.
```vbnet
    Private Function ValidationCallbackHandler() As Boolean
        Return True
    End Function
```

For information, the fetchSetting function is shown here, for more details see [User settings and the .xll.config file](user-settings-and-the-xllconfig-file.md)
```vbnet
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        Dim UserSettings As Collections.Specialized.NameValueCollection = Nothing
        Dim AddinAppSettings As Collections.Specialized.NameValueCollection = Nothing
        Try : UserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : End Try
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : End Try
        ' user specific settings are in UserSettings section in separate file
        If UserSettings(Key) Is Nothing Then
            If AddinAppSettings IsNot Nothing Then
                fetchSetting = AddinAppSettings(Key)
            Else
                fetchSetting = Nothing
            End If
        Else
            fetchSetting = UserSettings(Key)
        End If
        If fetchSetting Is Nothing Then fetchSetting = defaultValue
    End Function
```

