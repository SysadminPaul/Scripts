<#
.SYNOPSIS
    Using a CSV of all devices exported from Intune get a list of all configuration policies assigned to them then export as a CSV
.DESCRIPTION
    You can generate your own CSV as long as it has the headers Device ID, and Device Name.  The script relies on these to check Intune
#>

if (Get-Module -Name Microsoft.Graph.DeviceManagement){
    Write-Host -ForegroundColor Green "Microsoft.Graph.DeviceManagement Module already loaded.  Continuing"
}
elseif (Get-Module Microsoft.Graph.DeviceManagement -ListAvailable){
    Write-Host -ForegroundColor Cyan "Microsoft.Graph.DeviceManagement Module found.  Importing"
    Import-Module Microsoft.Graph.DeviceManagement
}
else{
    Write-Host -ForegroundColor Yellow "Microsoft.Graph.DeviceManagement Module not found.  Please install this first.  Script exiting"
    exit
}

Connect-MgGraph

# == NOTE - IF USING VSCODE THIS WILL OPEN BEHIND THE WINDOW ==
Write-Host -ForegroundColor Yellow "Dialog box opened for you to choose your CSV.  If you are using VSCode this will have loaded behind the main VSCode window"
Add-Type -AssemblyName System.Windows.Forms
$File = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath(‘Desktop’)
    Filter = "CSV (*.csv) | *.csv"
}
$File.ShowDialog()
$FilePath = $File.FileName
$devices = Import-Csv $FilePath

Write-Host -ForegroundColor Cyan "Working with $($devices.count) devices"

$MainOutput = @()
foreach ($device in $devices){
    Write-Host "Working with $($Device.PSObject.Properties["Device ID"].Value)"
    $deviceConfigPolicies = Get-MgDeviceManagementManagedDeviceConfigurationState -managedDeviceId $($Device.PSObject.Properties["Device ID"].Value)
    Write-Host "$($Device.PSObject.Properties["Device ID"].Value) | $($Device.PSObject.Properties["Device Name"].Value) | Has the following Configuration Policies - $($deviceConfigPolicies.displayName)"
    
    $MainOutput += [PSCustomObject]@{
        DeviceID = $($Device.PSObject.Properties["Device ID"].Value)
        DeviceName = $($Device.PSObject.Properties["Device Name"].Value)
        DeviceConfigPolicies = $($deviceConfigPolicies.displayName) -join ','
    }
}

# == NOTE - IF USING VSCODE THIS WILL OPEN BEHIND THE WINDOW ==
Write-Host -ForegroundColor Yellow "Dialog box opened for you to save your CSV.  If you are using VSCode this will have loaded behind the main VSCode window"
$SaveFile = New-Object System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath(‘Desktop’)
    Filter = "CSV (*.csv) | *.csv"
}
$SaveFile.ShowDialog()
$MainOutput | Export-CSV -NoTypeInformation $SaveFile.FileName