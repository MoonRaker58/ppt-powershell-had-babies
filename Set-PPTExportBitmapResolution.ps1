<#
.SYNOPSIS
    Set-PPTExportBitmapResolution is a way to change the resolution of PowerPoint's Export of slides as Bitmap images.
.DESCRIPTION
    Set-PPTExportBitmapResolution changes the resolution of PowerPoint's export of slides as Bitmap images, including the ability to reset to default. This affects the right click and "Save As Picture" option as well as the File > Save As option when the file format is a bitmap image type like jpeg, png, bmp.

    Parameter NewDpi has tab completion and validation of acceptable dpi values: 50, 96, 100, 150, 200, 250, 300

    Outputs a PSCustomObject representing a Registry Entry for a setting other than resetting to default.
.INPUTS
    System.Integer32
.PARAMETER NewDpi
    Specifies the new dpi. Tab completion based on list of acceptable values: 50, 96, 100, 150, 200, 250, 300
.PARAMETER ResetToDefault
    Switch to reset the dpi to default.
.OUTPUTS
    System.Management.Automation.PSCustomObject representing the RegistryEntry.
.EXAMPLE
    Set-PPTExportBitmapResolution
    Without parameters will set the dpi to the default.
    Returns nothing because this removes the Registry Entry.
.EXAMPLE
    Set-PPTExportBitmapResolution -ResetToDefault
    -ResetToDefault parameter will set the dpi to the default.
    Returns nothing because this removes the Registry Entry.
.EXAMPLE
    Set-PPTExportBitmapResolution -NewDpi 100
    NewDpi parameter will set the dpi to 100 and outputs:

    ExportBitmapResolution : 100
    PSPath                 : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options\
    PSParentPath           : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint
    PSChildName            : Options
    PSDrive                : HKCU
    PSProvider             : Microsoft.PowerShell.Core\Registry
.EXAMPLE
    Set-PPTExportBitmapResolution 100
    This will set the dpi to 100 and outputs:

    ExportBitmapResolution : 100
    PSPath                 : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options\
    PSParentPath           : Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint
    PSChildName            : Options
    PSDrive                : HKCU
    PSProvider             : Microsoft.PowerShell.Core\Registry
.LINK
    Background: https://docs.microsoft.com/en-us/office/troubleshoot/powerpoint/change-export-slide-resolution
#>
function Set-PPTExportBitmapResolution {
    [CmdletBinding()]
    [Alias('pptres')]
    [OutputType([string])]
    param(
        [ValidateSet(50, 96, 100, 150, 200, 250, 300)]
        [Alias('dpi','d')]
        [ArgumentCompletions(50, 96, 100, 150, 200, 250, 300)]
        [Parameter(ParameterSetName = "SetValue")]
        [int]
        $NewDpi,
        [Parameter(ParameterSetName = "Reset")]
        [Alias('reset','r')]
        [switch]
        $ResetToDefault
    )

    begin {
        # Find Office version so to maximize compatablity
        # If office is not installed, then nothing will be changed.
        $Keys = Get-Item -Path HKLM:\Software\RegisteredApplications | Select-Object -ExpandProperty property
        $Product = $Keys | Where-Object { $_ -Match "PowerPoint.Application." }
        $OfficeVersion = ($Product.Replace("PowerPoint.Application.", "") + ".0")


        # if (($null -eq $NewDpi) -or ($NewDpi -eq 96) -or ($NewDpi -notin 50, 96, 100, 150, 200, 250, 300)) {
        if (($null -eq $NewDpi) -or ($NewDpi -eq 96) -or ($NewDpi -eq 0)) {
            $ResetToDefault = $true
        }
    }

    process {
        if ($ResetToDefault) {
            $NewDpi = 96
            # Clear out any entry, ignore if entry doesn't exist
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Options\" -Name ExportBitmapResolution -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                ExportBitmapResolution = $NewDpi
                PSPath                 = ""
                PSParentPath           = ""
                PSChildName            = ""
                PSDrive                = ""
                PSProvider             = ""
            }

        }
        else {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Options\" -Name ExportBitmapResolution -Value $NewDpi -ErrorVariable SetEntryErrors -ErrorAction SilentlyContinue

            # This is to display message if the Office version can't be found in registry
            Write-Host $SetEntryErrors
        }
    }

    end {
        Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Options\" -Name ExportBitmapResolution -ErrorAction SilentlyContinue
    }
}