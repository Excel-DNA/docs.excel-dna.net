---
layout: page
title: "Modal dialog on new thread"
---
```csharp
new Thread(() => {
    var excelWindowThatIsTheOwner = new NativeWindow();

    excelWindowThatIsTheOwner.AssignHandle(new IntPtr(Application.Hwnd));

    // Show modal dialog (here: a message box)
    MessageBox.Show(owner: excelWindowThatIsTheOwner,
                    text: "I am a modal MessageBox.\r\nNow bring another application to the foreground and then try to bring excel back via the windows taskbar...");
}).Start();
```
