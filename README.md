# ppt-pwsh-had-babies
PowerPoint needs a better way to control Bitmap Export Resolution.

Microsoft's way is really painful.
https://docs.microsoft.com/en-us/office/troubleshoot/powerpoint/change-export-slide-resolution

I created a pure PowerShell way to change the dpi rather fast.

There is probably a better way but this way is mine.

Tested only on pwsh 7 on windows.

```
.SYNTAX
  Set-PPTExportBitmapResolution [[-NewDpi] <Int32>] [-ResetToDefault] [<CommonParameters>]
```

NewDpi parameter has tab completion based on list of acceptable values: 50, 96, 100, 150, 200, 250, 300

 ``
Set-PPTExportBitmapResolution -NewDpi 100
``
will set the dpi to 100 and output an unformated PSCustomObject that represents a Registry Entry:
```
    ExportBitmapResolution : 100
    PSPath                 : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options\
    PSParentPath           : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint
    PSChildName            : Options
    PSDrive                : HKCU
    PSProvider             : Microsoft.PowerShell.Core\Registry
```

``
Set-PPTExportBitmapResolution
``
Without parameters will set the dpi to the default.
Returns nothing because this removes the Registry Entry.

``
Set-PPTExportBitmapResolution -ResetToDefault
``
-ResetToDefault parameter will set the dpi to the default.
Returns nothing because this removes the Registry Entry.


If you're lucky and I feel like it, I may add a pester test someday.
Don't hold your breath.
