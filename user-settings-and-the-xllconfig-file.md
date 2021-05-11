---
layout: page
title: "User settings and the .xll.config file"
---

## Basic Usage

1. Make a file called `<TheAddInName>.xll.config` with this in:

```xml
<configuration>
  <appSettings>
    <add key="Test" value="Forty-two" />
  </appSettings>
</configuration>
```

2. In your project, add a reference to the System.Configuration assembly.

3. In your library add some function to access the settings:

```csharp
internal static string GetAppSetting(string key)
{
    object setting = System.Configuration.ConfigurationManager.AppSettings[key];
    if (setting == null)
    {
        return "!! INVALID KEY !!";
    }

    return setting.ToString();
}
```

4. If you run `ExcelDnaPack` to pack the add-in into a single file, the `.xll.config` file will automatically be packed too. At runtime, if a `.xll.config` file is present, it will be used. Otherwise the packed `.config` file will be used as the configuration for for the add-in's AppDomain.

## Advanced Topics

In the configuration file there are two ways to extend the settings to other files:

* by adding a UserSettings section in specifying it in the configSections element as a type of `System.Configuration.NameValueSectionHandler`
* by adding file names or paths to additional NameValueSection config files that enhance both the app settings and the user settings
	* the attribute name for the file name in the user settings is `configSource`,
	* the attribute name for app settings file is `file`

Below, there is an example of an app config (name it as described in [Basic Usage](#basic-usage)) that defines both app settings, including a separate AppSettings NameValueSection File and User Settings, including a separate UserSettings NameValueSection File as well.

```xml
<configuration>
  <configSections>
    <section name="UserSettings" type="System.Configuration.NameValueSectionHandler"/>
  </configSections>
  <UserSettings configSource="Your/Path/to/UserSettings.config">
     <add key="someSettingKey" value="someSettingValue"/>
  </UserSettings>
...
  <appSettings file="Your/Path/to/AppSettings.config">
    <add key="someOtherSettingKey" value="someOtherSettingValue"/>
  </appSettings>
...
</configuration>
```

The UserSettings and AppSettings NameValueSection files are just repetitions of the UserSettings or appSettings elements:

* content of `AppSettings.config`  

```xml
<appSettings>
    <add key="keyname0" value="3" />
    <add key="keyname1" value="False" />
</appSettings>
```

* content of `UserSettings.config`  

```xml
<UserSettings>
    <add key="keyname2" value="3" />
    <add key="keyname3" value="False" />
</UserSettings>
```

Having these three config files in place, you can then create a mechanism to have a central config file and user config files with either

* central config file overriding the user config files or
* user config files overriding the central config file,

depending on your needs.

Below, the second method is implemented in a VB.NET function that fetches a setting regardless whether it is found in the central or the user config, however the user config always has precedence. If nothing is found then the passed defaultValue is returned.
The class `NameValueCollection` is taken from `System.Collections.Specialized.NameValueCollection`

```vbnet
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        Dim AddinUserSettings As Collections.Specialized.NameValueCollection = Nothing
        Dim AddinAppSettings As Collections.Specialized.NameValueCollection = Nothing

        ' get the User Settings (in UserSettings section or in separate file), if available
        Try : AddinUserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : End Try
        ' get the App Settings (in AppSettings section or in separate file), if available
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : End Try

        ' UserSettings take precedence, only if setting is not available there then AddinAppSettings
        If AddinUserSettings(Key) Is Nothing Then
            If AddinAppSettings IsNot Nothing Then
                fetchSetting = AddinAppSettings(Key)
            Else
                fetchSetting = Nothing
            End If
        Else
            fetchSetting = AddinUserSettings(Key)
        End If
        ' if neither User nor AppSettings returned a value, take the defaultValue. Alternatively you could throw an error here.
        If fetchSetting Is Nothing Then fetchSetting = defaultValue
    End Function
```
