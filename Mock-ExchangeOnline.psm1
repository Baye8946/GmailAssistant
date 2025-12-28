# Mock-ExchangeOnline.psm1
function Mock-Connect-ExchangeOnline {
    param (
        [string]$UserPrincipalName
    )
    Write-Verbose "Mock: Connected to Exchange Online as $UserPrincipalName"
}

function Mock-Disconnect-ExchangeOnline {
    param (
        [switch]$Confirm
    )
    Write-Verbose "Mock: Disconnected from Exchange Online"
}

function Mock-New-TransportRule {
    param(
        [string]$Name,
        [string]$Mode,
        [string]$HeaderContainsMessageHeader,
        [string[]]$HeaderContainsWords,
        [string]$MoveToFolder,
        [bool]$DeleteMessage,
        [string]$RedirectMessageTo,
        [string]$Comments,
        [int]$Priority
    )
    Write-Verbose "Mock: Updating transport rule $Identity with conditions $($HeaderContainsWords -join ', ') and action $(if ($MoveToFolder) { "MoveToFolder: $MoveToFolder" } elseif ($DeleteMessage) { "Delete" } else {"Redirect: $RedirectMessageTo" })"
    return [PSCustomObject]@{
        Name = $Identity
        Mode = $Mode
        HeaderContainsWords = $HeaderContainsWords
    }
}

function Mock-Set-TransportRule {
    param (
        [string]$Identity,
        [string]$Mode,
        [string]$HeaderContainsMessageHeader,
        [string[]]$HeaderContainsWords,
        [string]$MoveToFolder,
        [bool]$DeleteMessage,
        [string]$RedirectMessageTo,
        [string]$Comments,
        [int]$Priority
    )
    Write-Verbose "Mock: Updating transport rule $Identity with conditions $($HeaderContainsWords -join ', ') and action $(if ($MoveToFolder) { "MoveToFolder: $MoveToFolder" } elseif ($DeleteMessage) { "Delete" } else { "Redirect: $RedirectMessageTo" })"
    return [PSCustomObject]@{
        Name = $Identity
        Mode = $Mode
        HeaderContainsWords = $HeaderContainsWords
    }
}

function Mock-Get-TransportRule {
    param (
        [string]$Identity
    )
    Write-Verbose "Mock: Checking for transport rule $Identity"
    #Simulate an existing rule
    return [PSCustomObject]@{
        Name = $Identity
        Mode = "Enforce"
        HeaderContainsWords = @("MeetingResponse", "MeetingDeclined")
    }
}

Export-ModuleMember -Function Mock-Connect-ExchangeOnline, Mock-Disconnect-ExchangeOnline, Mock-New-TransportRule, Mock-Set-TransportRule, Mock-Get-TransportRule