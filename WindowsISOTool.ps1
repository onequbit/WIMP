
function Debug-WIMPanic
{
    param
    (
        [parameter(Mandatory=$true)]
        [string] $ExceptionMessage
    )
    Write-Host "Mounted WIM exception - discarding..."
    Write-Verbose $ExceptionMessage -Verbose
    pause    
    Get-WindowsImage -Mounted | Dismount-WindowsImage -discard
    throw $ExceptionMessage
}


function DateTime
{
    return get-date -format "yyyyMMdd-HHmmss"
}

function LocalDirectory
{   
    $parent_dir = Split-Path -Path $env:SCRIPTPATH
    if ($parent_dir -eq $null -or $parent_dir -eq "")
    {
        throw "Could not extract local path"
    }    
    return $parent_dir
}

function IsAdmin
{
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}


function FileExtensionFilter
{
    param
    (
        [parameter(Mandatory,ValueFromPipeline)]
        [string]$Extension
    )
    process {
        $ucase = $Extension.ToUpper()
        $lcase = $Extension.ToLower()
        $filter = "$ucase files (*.$lcase)|*.$lcase"
        return $filter
    }
}

function FilePicker
{
    param
    (
        [parameter(Mandatory=$false)]
        [string]$Extension = $null,
        [parameter(Mandatory=$false)]
        [string[]]$Extensions = $null,
        [parameter(Mandatory=$false)]
        [string]$Title
        
    )
    $filter = ""
    if (-not $Extension -and -not $Extensions)
    {
        $filter = "All files|*.*"
    } 
    elseif ($Extension -and -not $Extensions) {
        $filter = $Extension | FileExtensionFilter        
    }
    elseif ($Extensions -and -not $Extension) {
        $filter_set = @()
        foreach ($ext in $Extensions)
        {
            $filter_set += $ext | FileExtensionFilter
        }
        $filter = $filter_set -Join "|"        
    }    
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $env:SystemDrive }
    $FileBrowser.ShowHelp = ($Host.name -eq "ConsoleHost")
    $FileBrowser.Filter = $filter
    if ($Title)
    {
        $FileBrowser.Title = $Title
        Write-Host $Title
    }
    [void]$FileBrowser.ShowDialog()
    return $FileBrowser.FileName.ToString()
}

function FolderPicker
{
    param
    (
        [parameter(Mandatory=$false)]
        [string] $Description = "Select a folder",
        [parameter(Mandatory=$false)]
        [string] $InitialPath = "MyComputer"
        

    )
    Write-Host $Description
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = $Description
    $FolderBrowser.RootFolder = $InitialPath
    $FolderBrowser.SelectedPath = ((get-psdrive).Root | ? {$_ -like '*:*'})[0]
    [void] $FolderBrowser.ShowDialog()
    return $FolderBrowser.SelectedPath.ToString()
}

function New-Directory
{    
    param
    (
        [parameter(Mandatory=$true)]
        [string] $Parent,
        [parameter(Mandatory=$true)]
        [string] $Name
    )
    $new_path = "$Parent\$Name"
    try {
        New-Item -Path $Parent -Name $Name -ItemType Directory | Out-Null
        attrib -r $new_path
        $is_valid = Test-Path $new_path
        if (!$is_valid)
        {        
            throw "Failed to create new directory $new_path"
        }
        
        return [string]$new_path
    }
    catch {
        return ""
    }
}

function Get-RequiredFile
{
    param(
        [Parameter(Mandatory=$true)]
        [string] $Path,        
        [Parameter(Mandatory=$true)]
        [string] $Filename
    ) 
    Write-Host "Checking $Path for $Filename"
    $FilePath = (Get-ChildItem -Path $Path -Include $Filename -Recurse -ErrorAction SilentlyContinue)
    if (($FilePath.Length -eq 0))
    {        
        throw "The following file: '$Filename' is required in the same folder as this script"
    }    
    return $FilePath[0].FullName
}

function Copy-DiscToDraft
{
    param(
        [parameter(Mandatory)] [string] $ISODrive,
        [parameter(Mandatory)] [string] $DraftTarget
    )

    if (!(Test-Path $DraftTarget))
    {
        throw "Invalid path: $DraftTarget"
    }
    try {
        Copy-Item -Path "$ISODrive\*" -Destination $DraftTarget -Container -Recurse -ErrorAction Stop

        Write-TextBanner -Text "$ISODrive\* copied successfully to $DraftTarget"
    }
    catch {
        Write-Verbose $_ -Verbose
        Write-Host "ISOContent: '$ISOContent'"
        Write-Host "DraftTarget: '$DraftTarget'"
        throw "failed to copy ISO to draft folder"
    }
}

function Add-CustomFile
{
    param
    (
        [parameter(Mandatory=$true)]
        [string] $Source,
        [parameter(Mandatory=$true)]
        [string] $TargetFolder,
        [parameter(Mandatory=$true)]
        [string] $TargetFileName        
    )
    $DestinationFilePath = "$TargetFolder\$TargetFileName"
    try {
        if (!(Test-Path $TargetFolder))
        {
            Write-Host "$TargetFolder not found, creating"
            mkdir $TargetFolder | Out-Null
        }          
    } catch {
        Write-TextBanner -Text "Error! Unable to use $TargetFolder to deposit $SourceFilePath as $TargetFileName"
    }
    
    Write-Host "Adding '$Source' as '$DestinationFilePath'"
    try {
        Copy-Item -Path $Source -Destination $DestinationFilePath -force
    } catch {
        Write-Host "Failed to copy $Source to $DestinationFilePath"
        Write-Host "Source: '$Source', TargetFolder: '$TargetFolder', TargetFileName: '$TargetFileName', New File: '$DestinationFilePath'"
        Write-Verbose $_ -Verbose
    }    
}
function ShowBalloonTip
{
    param(        
        [Parameter(Mandatory=$true)]
        [string] $Title,
        [Parameter(Mandatory=$true)]
        [string] $Text,
        [Parameter(Mandatory=$false)]
        [int] $Seconds = 15
    )
    $Miliseconds = $Seconds * 1000
    Add-Type -AssemblyName System.Windows.Forms 
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info 
    $balloon.BalloonTipText = "$Text"
    $balloon.BalloonTipTitle = "$Title" 
    $balloon.Visible = $true 
    $balloon.ShowBalloonTip($Miliseconds)
    $balloon.Dispose()
}

function Mount-WIMFile
{

    param(
        [parameter(Mandatory)] [string] $FilePath,
        [parameter(Mandatory)] [string] $MountPath
    )

    if (!(Test-Path $MountPath))
    {
        throw "Invalid path: $MountPath"
    }

    try {    
        attrib -r $MountPath            
        Mount-WindowsImage -Path $MountPath -ImagePath $FilePath -Name "Windows 10 Pro" | Out-Null

        Write-TextBanner -Text "$FilePath mounted successfully at $MountPath"

        ShowBalloonTip -Title "Script Status:" -Text "$FilePath mounted successfully at $MountPath" -Seconds 45

    } catch {
        Write-Verbose $_ -Verbose
        throw "Failed to mount WIM image $FilePath onto $MountPath"
    }
}

function Get-WindowsVersionCode
{
    param(        
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [object] $BuildNumber
    )

    process {
        $VersionCode = @{}
        $VersionCode["19043"] = "21H1"
        $VersionCode["19042"] = "20H2"
        $VersionCode["19041"] = "2004"
        $VersionCode["18363"] = "1909"
        $VersionCode["18362"] = "1903"
        $VersionCode["17763"] = "1809"
        $VersionCode["17134"] = "1803"
        $VersionCode["16299"] = "1709"
        $VersionCode["15063"] = "1703"
        $VersionCode["14393"] = "1607"
        $VersionCode["10586"] = "1511"
        
        if ($BuildNumber.Trim() -ne "")    
        {
            $BuildNumber = ($BuildNumber.Split('\n').Trim() | Where-Object { $_ -ne "BuildNumber" }) -Join ""
            $number = 0
            try {
                $number = [int] ($BuildNumber.Trim())
            } catch {
                return
            }        
            return $VersionCode["$number"]        
        }
    }
}

function Get-WIMBuildNumber
{
    param(        
        [Parameter(Mandatory=$true)]
        [string] $WIMFilePath
    )
    $wim_image_info = Get-WindowsImage -ImagePath "$WIMFilePath" -Name "Windows 10 Pro"    
    return $wim_image_info.Version.Split('.')[2]
}

function Remove-MSBloatWare
{
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string] $WIMMountPath
    )

    $packages_to_remove = @(
        "Microsoft.BingWeather",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.MixedReality.Portal",
        "Microsoft.People",
        "Microsoft.SkypeApp",
        "Microsoft.Wallet",
        "Microsoft.Windows.Photos",
        "Microsoft.WindowsAlarms",
        "Microsoft.WindowsCamera",
        "microsoft.windowscommunicationsapps",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.WindowsMaps",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxApp",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo"
    )
    
    $wim_mount_path = (Get-WindowsImage -Mounted).Path
    $installed_packages = Get-AppxProvisionedPackage -Path $wim_mount_path

    foreach ($package_name in $packages_to_remove)
    {
        try {
            $package = $installed_packages | Where-Object {$_.DisplayName -eq $package_name}
            if ($null -ne $package)
            {
                Remove-AppxProvisionedPackage -Path $wim_mount_path -PackageName $package.PackageName | Out-Null
                Write-Host "$package_name removed from $wim_mount_path"
            }            
        } catch {
            Write-Host "Failed to remove package: $package"
            Write-Verbose $_ -Verbose            
        }
    }    
}

function Add-ThirdPartyInstallers
{
    param(
        [parameter(Mandatory)] [string] $SourceFolder,
        [parameter(Mandatory)] [string] $TargetFolder
    )

    if ($SourceFolder)
    {
        # $TargetFolder = (New-Directory -Parent ((Get-WindowsImage -Mounted).Path) -Name "Installers")
        Write-Host "Copying $SourceFolder into $TargetFolder..."
        try {
            #$exclude = Get-ChildItem -recurse $TargetFolder
            $err = @()
            Copy-Item -Path "$SourceFolder\*" -Destination $TargetFolder -Container -Recurse -Force -ErrorVariable +err -ErrorAction SilentlyContinue
        }
        catch {
            Write-Verbose $_ -Verbose            
        }
        if ($err.Count -gt 0)
        {
            Write-Host "Copy errors:"
            Write-Host $err
        }
        Write-TextBanner -Text "...3rd-party Installers added"
    } else 
    {
        Write-Host "No installer folder selected... skipping"
    }   
}

function Add-UnattendedAnswerFile
{
    param(                
        [parameter(Mandatory)][string] $DraftTarget,
        [parameter(Mandatory)][string] $UnattendPath
    )
    if (!(Test-Path $DraftTarget))
    {
        throw "Invalid path: $DraftTarget"
    }

    try {
        Add-CustomFile -Source $UnattendPath -TargetFolder $DraftTarget -TargetFileName "autounattend.xml"
        
        $PantherPath = "$WIMMountPath\Windows\Panther"
        Add-CustomFile -Source $UnattendPath -TargetFolder $PantherPath -TargetFileName "unattend.xml"

        Write-TextBanner -Text "$UnattendPath added to image"
        
    } catch {
        Write-Verbose $_ -Verbose
        Debug-WIMPanic -ExceptionMessage "Adding files to the image ran into a problem."
    }
}


function Save-WIMImage
{
    try {
        $mounted_image = Get-WindowsImage -Mounted 
        $mounted_image | Dismount-WindowsImage -Save -CheckIntegrity | Out-Null
        Write-Host "$($mounted_image.ImagePath) dismounted safely"
        ShowBalloonTip -Title "Script Status:" -Text "$($mounted_image.ImagePath) dismounted safely" -Seconds 20
    }
    catch {
        Write-Verbose $_ -Verbose
        Debug-WIMPanic -ExceptionMessage "Failed to dismount WIM image"    
    }
}

function Write-FolderToISO
{
    param
    (
        [parameter(Mandatory)]
        [string]$Folder,
        [parameter(Mandatory)]
        [string]$ISOFilePath        
    )
    try {        
        $iso_writer_path = Get-RequiredFile -Path ".\" -Filename "oscdimg.exe"        
        $boot_image_path = Join-Path -Path $Folder -ChildPath "efi\microsoft\boot\efisys.bin"    
        $create_ISO = "$iso_writer_path -u2 -m -b$boot_image_path $Folder $ISOFilePath"
        Write-Host "Invoking: '$create_ISO'"
        Invoke-Expression -Command $create_ISO | Out-Null

    }
    catch {
        Write-Verbose $_ -Verbose
        throw "Failed to dismount WIM image"
    }    
}

function Write-TextBanner
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $Text
    )
    $newline = "`r`n"
    [int]$length = $Text.Length
    $inner_text = "|   " + $Text + "   |$newline"    
    $bar = "+---" + ("-" * $length) + "---+$newline"
    $inner_space = "|   " + (" " * $length) + "   |$newline"
    $box = $bar + $inner_space + $inner_text + $inner_space + $bar
    Write-Host $box
}


########################################################################################################

### This script must be run as administrator, and will now attempt to elevate itself with those credentials.
# *** 
# *** The code for the script reinvoking itself can't be put inside of a function, 
# *** doing so prevents it from behaving as expected.
# *** 
$env:SCRIPTPATH = $MyInvocation.MyCommand.Path

$these_args = $args -join " "
$reinvocation = "-windowstyle normal -noexit -file $env:SCRIPTPATH $these_args"
if (!(IsAdmin))
{
    Start-Process -filepath "powershell" -ArgumentList $reinvocation -Verb RunAs
    return
}

try {
    $localpath = LocalDirectory    
    Set-Location "$localpath"
    $title = "Windows Installation Media Preparer (WIMP)"
    $Host.UI.RawUI.WindowTitle = $title
    Write-TextBanner -Text $title
}
catch {
    Write-Verbose $_ -Verbose
    throw
}

########################################################################################################

### Verify that the following files are present in the local directory (or subdirectory):

Write-Host "Verifying required files(s)"

$iso_writer_path = Get-RequiredFile -Path ".\" -Filename "oscdimg.exe"

Write-TextBanner -Text "Required file verified"

########################################################################################################

### Get the unattended answer file

Write-Host "Select the unattended answer file"

try {
    $DesiredUnattendFile = FilePicker -Extension "xml" -Title "Select the Unattended Answer File"
    if (!($DesiredUnattendFile))
    {
        throw "Unattend file not selected"
    }
    Write-TextBanner -Text "Unattend file: $DesiredUnattendFile"        
    
}
catch {
    Write-Verbose $_ -Verbose
    throw "failed to identify unattend file"
}

########################################################################################################

### Get the input file name for the ISO to be customized
Write-Host "Get the input files for the ISO to be customized"

try {
    $SourceISOFileName = FilePicker -Extension "iso" -Title "Select the ISO to extract information"
    if (!($SourceISOFileName))
    {
        throw "Source ISO not selected"
    }
    Write-TextBanner -Text "Source ISO: $SourceISOFileName"        
    
}
catch {
    Write-Verbose $_ -Verbose
    throw "failed to identify source ISO to be customized"
}


########################################################################################################

### Mount ISO To Be Copied and show the relevant Download site for Updates
Write-Host "Mount ISO $SourceISOFileName stage."

$ISODrive = ''
try {
    if (!(Test-Path $SourceISOFileName))
    {
        Write-Host "Path: [$SourceISOFileName] is invalid!"
        throw "Failed to mount ISO $SourceISOFileName"
    }
    Write-Host "Mounting $SourceISOFileName..."
    Mount-DiskImage -ImagePath $SourceISOFileName | Out-Null
    $ISODrive = (Get-DiskImage -ImagePath "$SourceISOFileName" | Get-Volume).DriveLetter + ":"
    $WIMFilePath = "$ISODrive\sources\install.wim"
    $BuildNumber = Get-WIMBuildNumber -WIMFilePath $WIMFilePath
    $VersionCode = Get-WindowsVersionCode -BuildNumber $BuildNumber

}
catch {
    Write-Verbose $_ -Verbose
    throw "ISO Mount failed for $SourceISOFileName"
}

Write-TextBanner -Text "Mounted $SourceISOFileName to $ISODrive"

########################################################################################################

### Create Working folder and subfolders

Write-Host "Create Working folder and subfolders"

$WorkingDir = ""
try {    
    $WorkingDir = FolderPicker -Description "Select a working folder for the ISO draft and Image Mount - folder must be empty"
    if (!($WorkingDir))
    {
        throw "Working Directory not selected."
    }
    Write-TextBanner -Text "Working folder $WorkingDir selected."    

    $InstallersFolder = FolderPicker -Description "Select the folder containing additional application installers to add to the image"

}
catch {
    Write-Verbose $_ -Verbose
    throw "failed to select working folder"
}

if ((Get-ChildItem $WorkingDir).Length -gt 0)
{
    throw "$WorkingDir must be empty"
}

$DraftTarget = New-Directory -Parent $WorkingDir -Name "draft"
Write-Host "Draft ISO folder $DraftTarget created."
$WIMMountPath = New-Directory -Parent $WorkingDir -Name "mount"    
Write-Host "Image mount-point $WIMMountPath created."

########################################################################################################


### Copy ISO To Draft Folder

Write-Host "Copying $ISODrive to $DraftTarget"
Copy-DiscToDraft -ISODrive $ISODrive -DraftTarget $DraftTarget

########################################################################################################

### Dismount ISO - contents copied, we can dismount it

Write-Host "Dismount ISO $SourceISOFileName"

try {
    Dismount-DiskImage -ImagePath $SourceISOFileName | Out-Null
}
catch {
    Write-Verbose $_ -Verbose
    throw "failed to dismount ISO"
}
Write-TextBanner -Text "ISO dismounted"

########################################################################################################

### Extract WIM File and image as needed

Write-TextBanner -Text "Extracting information from $WIMFilePath"

$WIMFilePath = "$DraftTarget\sources\install.wim" 
if (!(Test-Path $WIMFilePath))
{
    throw "Invalid path: $WIMFilePath"
}

$ImageInfo = ''
try {
    Write-Host "Unsetting readonly attribute for $WIMFilePath"
    attrib -r $WIMFilePath
    Write-Host "Reading $WIMFilePath for available Images" 
    $ImageInfo = Get-WindowsImage -ImagePath "$WIMFilePath"
    Write-Host "Found 'Windows 10 Pro' Image in $WIMFilePath"
}
catch {
    Write-Verbose $_ -Verbose
    throw "Get-WindowsImage command failed"
}

Write-TextBanner -Text "Isolating Image from Multi-image ISO"

if ($ImageInfo.Length -ne 1)
{
    # We need to rename the WIM file so that the desired image can be extracted.
    # This will reduce the size of the WIM file and simplify the process.
    try {            
        $renamed_WIM_path = "$DraftTarget\sources\install_m.wim"
        Rename-Item -Path $WIMFilePath -NewName $renamed_WIM_path
        Export-WindowsImage -SourceImagePath $renamed_WIM_path -DestinationImagePath $WIMFilePath -SourceName "Windows 10 Pro" | Out-Null                    
        Remove-Item -Path $renamed_WIM_path -Force -Confirm:$false
    }
    catch {
        Write-Verbose $_ -Verbose
        throw "failed to isolate Image from Multi-Image ISO"
    }
}

########################################################################################################

### Mount WIM File

Write-Host "Preparing to mount Windows 10 Pro image $WIMFilePath at $WIMMountPath"

Mount-WIMFile -FilePath $WIMFilePath -MountPath $WIMMountPath

########################################################################################################

### Remove Bloatware Packages

Write-Host "Removing Provisioned Packages"

Remove-MSBloatWare -WIMMountPath ((Get-WindowsImage -Mounted).Path)

Write-TextBanner -Text "...provisioned packages removed successfully"

########################################################################################################

### Add 3rd-party installation media and update files

$TargetFolder = (New-Directory -Parent ((Get-WindowsImage -Mounted).Path) -Name "Installers")

Write-TextBanner -Text "Adding 3rd-party Installers..."
Add-ThirdPartyInstallers -SourceFolder $InstallersFolder -TargetFolder $TargetFolder

# Write-TextBanner -Text "Retrieving Updates for this Windows Build"
# Get-WindowsUpdates -VersionCode $VersionCode -DownloadPath $TargetFolder


########################################################################################################

### Apply Unattend File

Write-Host "Adding Unattended Answer file to image"

Add-UnattendedAnswerFile -DraftTarget $DraftTarget -UnattendPath $DesiredUnattendFile

########################################################################################################

### Dismount WIM File

Write-Host "Dismounting and saving changes to $WIMFilePath..."

Save-WIMImage

Write-TextBanner -Text "$WIMFilePath dismounted and saved."


########################################################################################################

### Writing ISO

Write-Host "Writing $OutputISO..."

Write-FolderToISO -Folder $DraftTarget -ISOFilePath "$env:userprofile\Desktop\WindowsInstall_$(DateTime).iso"

Write-TextBanner -Text "...created $OutputISO successfully."

########################################################################################################

### Show output

Write-Host "opening file $OutputFileName"
$logfile = "$env:windir\Logs\DISM\Dism.log"

ShowBalloonTip -Title "Script Status:" -Text "Log file: $logfile" -Seconds 30
ShowBalloonTip -Title "Script Status:" -Text "ISO created: $OutputISO" -Seconds 30
@"
#####################
#                   #
#     Job's done!   #
#                   #
#####################
"@

Start-Process "$env:userprofile\Desktop\"


########################################################################################################
########################################################################################################
########################################################################################################
