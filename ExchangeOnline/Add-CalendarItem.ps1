<#
    .SYNOPSIS
        Adds calendar items to staff calendars
    .EXAMPLE
        Add-CalendarItems.ps1 -CalendarItems (Import-CSV .MyCSVFile.csv)
        Run the script using a specific CSV file
    .EXAMPLE
        Add-CalendarItems.ps1 -NoDisconnect
        Run the script but do not disconnect from Microsoft Graph after completion
    .NOTES
        Scopes required were found using Find-MgGraphCommand -uri "/devices".  The output of this tells you for the specified URI which permissions are required.
    .PARAMETER CalendarItems
        Must be result of an Import-CSV.  File must be in the following format:

        StartDate,EndDate,Subject,StartDateMinusOne,EndDatePlusOne
        2022-12-15T00:00:00,2022-12-16T00:00:00,CSV Test Item2022-12-14T00:00:00,2022-12-17T00:00:00
    .PARAMETER Group
        ObjectID of the AzureAD Group you want to perform the calendar add to.
    .PARAMETER ClientID
        ClientID of the Azure App
    .PARAMETER TenantId
        TenantID of the Azure tenant the app has been created in
    .PARAMETER NoDisconnect
        Do not disconnect from Microsoft Graph after completion
    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate used as the secret for authenticating against the app.  Use either this or the Certificate, not both.
    .PARAMETER Certificate
        The certificate used as the secret for authentication against the app.  Use either this or the CertificateThumbprint, not both.
    #>

    [CmdletBinding()]
    param(
        [Parameter()]
        $CalendarItems,
        [Parameter()]
        [string] $Group = "",
        [Parameter()]
        [switch] $NoDisconnect,
        [Parameter()]
        [string] $ClientId = "",
        [Parameter()]
        [string] $TenantId = "",
        [Parameter()]
        [string] $CertificateThumbprint,
        [Parameter()]
        $Certificate
    )
        
    BEGIN {
        if (!$CalendarItems) { $CalendarItems = Import-CSV $PSScriptRoot/CalendarItems-entries.csv }
        
        Write-Host -ForegroundColor Cyan "Prerequisite check starting"
        if ( -not (Get-Module Microsoft.Graph -ListAvailable)) { 
            Write-Host -ForegroundColor Yellow "Microsoft.Graph PowerShell Module not found.  Attempting install..."
            Install-Module Microsoft.Graph -Confirm:$false -Force
            Import-Module Microsoft.Graph.Calendar
            Import-Module Microsoft.Graph.Groups
        }
        else {
            Write-Host -ForegroundColor Green "Microsoft.Graph PowerShell Module found.  Loading..."
            Import-Module Microsoft.Graph.Calendar
            Import-Module Microsoft.Graph.Groups
        }
        if (-not(((get-MgContext).AuthType -contains "AppOnly") -and ((Get-MgContext).AppName -contains "PowerShell Exchange"))) {
            Write-Host -ForegroundColor Cyan "Microsoft Graph connection either does not exist or does not have correct scopes.  Connecting to Microsoft Graph..."
            try {
                if ($CertificateThumbprint) {
                    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
                }
                else {
                    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $Certificate    
                }
            }
            catch {
                Write-Host -ForegroundColor Red "Unknown error occured when attempting to connect to graph.  Exiting"
                exit
            }
        }
        else {
            Write-Host -ForegroundColor Cyan "Microsoft Graph connection already extablished.  Continuing..."
        }
        
        $MgGroup = Get-MgGroupMember -GroupId $Group -All
        
    }
             
    PROCESS {
        Write-Host -ForegroundColor Cyan "Scipt process starting"
        
        foreach ($Appt in $CalendarItems) {
            $params = @{
                Subject       = $Appt.Subject
                Start         = @{
                    DateTime = $Appt.StartDate
                    TimeZone = "GMT Standard Time"
                }
                End           = @{
                    DateTime = $Appt.EndDate
                    TimeZone = "GMT Standard Time"
                }
                StartMinusOne = @{
                    DateTime = $Appt.StartDateMinusOne
                    TimeZone = "GMT Standard Time"
                }
                EndPlusOne    = @{
                    DateTime = $Appt.EndDatePlusOne
                    TimeZone = "GMT Standard Time"
                }
            }
                
            foreach ($User in $MgGroup) {
                $ParamsStart = $Params.start.DateTime
                $ParamsEnd = $Params.end.DateTime
                $ParamsStartMinusOne = $Params.StartMinusOne.DateTime
                $ParamsEndPlusOne = $Params.EndPlusOne.DateTime
                $ParamsSubject = $Params.Subject
                $User = Get-MGUser -UserId $User.id
                $Cal = Get-MgUserCalendar -UserId $User.Id | Where Name -eq 'Calendar'
                Write-Verbose "Working on $($User.UserPrincipalName) with calendar id $($Cal.id) called $($Cal.name)"
        
                if (Get-MgUserCalendarEvent -UserID $User.id -CalendarID $Cal.Id -Filter "start/datetime ge '$ParamsStartMinusOne' and end/datetime le '$ParamsEndPlusOne' and subject eq '$ParamsSubject'") {
                    Write-Host -ForegroundColor Cyan "Event: $ParamsSubject with start date: $ParamsStart already exists in $($User.UserPrincipalName) Calendar.  Skipping"
                    }
                else {
                    Write-Host -ForegroundColor Green "No Event: $ParamsSubject with start date: $ParamsStart found for $($User.UserPrincipalName).  Creating"
                    New-MgUserCalendarEvent -UserID $User.id -CalendarID $Cal.Id -IsAllDay -Subject $params.Subject -Start $params.Start -End $params.End -ShowAs Free -IsReminderOn:$false | Out-Null
                }
            }
        }
    }
        
    END {
        if (-not $NoDisconnect) {
            Write-Host -ForegroundColor Cyan "Signing out of Microsoft graph"
            Disconnect-MgGraph | Out-Null
        }
        
        Write-Host -ForegroundColor Cyan "Script finished"
    }