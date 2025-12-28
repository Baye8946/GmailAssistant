<#
.SYNOPSIS
    Automates creation or updating of an Exchange Online transport rule to filter Calendar Accept/Decline responses.
.DESCRIPTION
    Creates or updates a transport rule to filter Calendar Accept, Decline, or both event types, moving them to a specified folder or deleting them.
.PARAMETER FilterAccept
    Filter Calendar Accept responses.
.PARAMETER FilterDecline
    Filter Calendar Decline responses.
.PARAMETER Action
    Action to take: MoveToFolder, Delete, or Redirect.
.PARAMETER TargetFolder
    Folder to move filtered emails (if Action is MoveToFolder).
.PARAMETER RuleName
    Name of the transport rule (default: CalendarResponseFilter).
.EXAMPLE
    .\FilterCalendarResponses.ps1 -FilterAccept -FilterDecline -Action MoveToFolder -TargetFolder "Calendar Responses" -RuleName "CalendarResponseFilter" -Verbose
#>

[CmdletBinding()]
param (
    [Parameter()]
    [switch]$FilterAccept,
    [Parameter()]
    [switch]$FilterDecline,
    [Parameter(Mandatory = $true)]
    [ValidateSet("MoveToFolder", "Delete", "Redirect")]
    [string]$Action,
    [Parameter()]
    [string]$TargetFolder,
    [Parameter()]
    [string]$RuleName = "CalendarResponseFilter",
    [Parameter()]
    [switch]$UseMock
)

begin {
    if ($UseMock) {
        Write-Verbose "Using Mock Exchange Online cmdlets"
        Import-Module -Name ".\Mock-ExchangeOnline.psm1" -Force
        Mock-Connect-ExchangeOnline -UserPrincipalName "admin@localhost"
    } else {
        Write-Verbose "Connecting to Exchange Online..."
        try {
            Connect-ExchangeOnline -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
            return
        }
    }
}

process {
    Write-Verbose "Checking for existing transport rule: $RuleName"
    $existingRule = if ($UseMock) { Mock-Get-TransportRule -Identity $RuleName } else { Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue }

    $conditions = @()
    if ($FilterAccept) { $conditions += "MeetingResponse" }
    if ($FilterDecline) { $conditions += "MeetingDeclined " }
    if (-not $conditions) {
        Write-Error "At least one of -FilterAccept or -FilterDecline must be specified."
        return
    }

    $ruleParams = @{
        Name = $RuleName
        Priority = 0
        Mode = "Enforce"
        HeaderContainsMessageHeader = "Message-Type"
        HeaderContainsWords = $conditions
        Comments = "Filters Calendar Accept/Decline responses. Created/Updated: $(Get-Date)"
    }

    if ($Action -eq "MoveToFolder") {
        if (-not $TargetFolder) {
            Write-Error "TargetFolder is required when Action is MoveToFolder."
            return
        }
        $ruleParams["MoveToFolder"] = $TargetFolder
    } elseif ($Action -eq "Delete") {
        $ruleParams["DeleteMessage"] = $true
    } elseif ($Action -eq "Redirect") {
        $ruleParams["RedirectMessageTo"] = $TargetFolder
    }

    try {
        if ($existingRule) {
            Write-Verbose "Updating existing rule: $RuleName"
            if ($UseMock) {
                Mock-Set-TransportRule -Identity $RuleName @ruleParams
            } else {
                Set-TransportRule -Identity $RuleName @ruleParams -ErrorAction Stop
            }
            Write-Host "Updated transport rule: $RuleName"
        } else {
            Write-Verbose "Creating new transport rule: $RuleName"
            if ($UseMock) {
                Mock-New-TransportRule @ruleParams
            } else {
                New-TransportRule @ruleParams -ErrorAction Stop
            }
            Write-Host "Created transport rule: $RuleName"
        }
    }
    catch {
        Write-Error "Failed to create/update transport rule: $($_.Exception.Message)" #Do this if a terminating exception happens#
    }
}


end {
    if ($UseMock) {
        Mock-Disconnect-ExchangeOnline
    } else {
        Write-Verbose "Disconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false
    }
}