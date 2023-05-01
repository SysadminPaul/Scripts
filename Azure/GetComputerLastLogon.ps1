    <#
    .SYNOPSIS
        Obtains details from the last logon to a specified computer name from Azure audit logs.  The computer must be Azure joined or Azure Hybrid joined for logins to appear.
    .EXAMPLE
        Get-GraphComputerLastLogin -Computername Computer01
        Show the last login for the computer "Computer01"
    .EXAMPLE
        Get-GraphComputerLastLogin -Computername Computer01 -Last30Days
        Show the last 30 days worth of logins for the computer "Computer01".  This will be output in a grid view for sorting/filtering.
    .NOTES
        Scopes required were found using Find-MgGraphCommand -uri "/devices".  The output of this tells you for the specified URI which permissions are required.
    .PARAMETER ComputerName
        The name of the computer you are searching logs for
    .PARAMETER Last30Days
        Display the last 30 days worth of logins.  By default only the last login is shown.
    .PARAMETER NoDisconnect
        Do not disconnect from Microsoft Graph after completion
    #>
     
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0 )]
        [string] $ComputerName,
        [Parameter()]
        [switch] $Last30Days,
        [Parameter()]
        [switch] $NoDisconnect,
        [Parameter()]
        [string] $ClientId,
        [Parameter()]
        [string] $TenantId,
        [Parameter()]
        [string] $CertificateThumbprint
    )
     
    BEGIN {
        Write-Host -ForegroundColor Cyan "Prerequisite check starting"
        if ( -not (Get-Module Microsoft.Graph -ListAvailable)) { 
            Write-Host -ForegroundColor Yellow "Microsoft.Graph PowerShell Module not found.  Attempting install..."
            Install-Module Microsoft.Graph -Scope CurrentUser
            Import-Module Microsoft.Graph.AuditLog
        }
        else {
            Write-Host -ForegroundColor Green "Microsoft.Graph PowerShell Module found.  Loading..."
            Import-Module Microsoft.Graph.Reports
        }
        
        if (-not(((get-MgContext).scopes -contains "AuditLog.Read.All") -and ((Get-MgContext).Scopes -contains "Directory.Read.All"))){
            Write-Host -ForegroundColor Cyan "Microsoft Graph connection either does not exist or does not have correct scopes.  Connecting to Microsoft Graph"
            
            if (($ClientId) -or ($TenantId) -or ($CertificateThumbprint) ){
                Write-Host -ForegroundColor Cyan "Connecting as an application"
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
            }
            else{
                Write-Host -ForegroundColor Cyan "Connecting as a user"
                Connect-MgGraph -Scopes "AuditLog.Read.All, Directory.Read.All" | Out-Null
            }
        }
        else{
            Write-Host -ForegroundColor Cyan "Valid Microsoft Graph connection already established."
        }
    }
     
    PROCESS {
        Write-Host -ForegroundColor Cyan "Scipt process starting"
        $Signins = Get-MgAuditLogSignIn -Filter "deviceDetail/displayName eq '$ComputerName' and appDisplayName eq 'Windows Sign In'"
        $filterDate = (Get-Date).AddDays(-30).Date
        $FilteredResults = $signins | Where-Object {$_.CreatedDateTime -ge $filterDate}
        if(-not$Last30Days){
            $FilteredResults | Select-Object @{N="ComputerName";E={$_.DeviceDetail.DisplayName}}, UserPrincipalName, CreatedDateTime -First 1 | Format-Table
        } 
        else{
            $FilteredResults | Select-Object @{N="ComputerName";E={$_.DeviceDetail.DisplayName}}, UserPrincipalName, CreatedDateTime | Out-GridView
        }
    }
     
    END {
        if(-not$NoDisconnect){
        Write-Host -ForegroundColor Cyan "Script finished.  Signing out of Microsoft graph"
        Disconnect-MgGraph | Out-Null
        }
    }
