# GmailDailySummary - Daily email summary with categorization
# Requirements: client_secret.json and gmail_token.json from previously written scripts

# Load credentials
$ClientJson = Get-Content -Path "C:\SandboxShare\client_secret.json" -Raw | ConvertFrom-Json
$ClientID = $ClientJson.installed.client_id
$ClientSecret = $ClientJson.installed.client_secret

$TokenJson = Get-Content -Path "C:\SandboxShare\gmail_token.json" -Raw | ConvertFrom-Json
$RefreshToken = $TokenJson.refresh_token

# Refresh access token
try {
    $RefreshBody = @{
    client_id = $ClientID
    client_secret = $ClientSecret
    refresh_token = $RefreshToken
    grant_type = "refresh_token"
    }

$RefreshResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $RefreshBody -ContentType "application/x-www-form-urlencoded"
$AccessToken = $RefreshResponse.access_token

} catch {
    Write-Error "Failed to refresh token: $($_.Exception.Message)"
    return
}

$Headers = @{
    Authorization = "Bearer $AccessToken"
}

# Correct endpoint
$ApiUrl = "https://www.googleapis.com/g,ail/v1/users/me/messages"

# Get emails from last 24 hours
$Yesterday = (Get-Date).AddDays(-1).ToUniversalTime().ToString("yyyy/MM/dd")
$Query = "after:$Yesterday" 

try {
$MessagesResponse = Invoke-RestMethod -Uri "https://www.googleapis.com/gmail/v1/users/me/messages?q=$Query&maxResults=50" -Headers $Headers
} catch {
    Write-Error "Failed to fetch emails: $($_.Exception.Message)"
    Write-Host "Check internet connection or DNS (try pinging www.googleapis.com)"
    return
}

if (-not $MessagesResponse.messages) {
    Write-Host "No new emails in the last 24 hours." -ForegroundColor Green
}

$Emails = @()

foreach ($msg in $MessagesResponse.messages) {
    $EmailDetail = Invoke-RestMethod -Uri "https://www.googleapis.com/gmail/v1/users/me/messages/$($msg.id)?format=metadata" -Headers $Headers

    $HeadersObj = $EmailDetail.payload.headers
    $From = ($HeadersObj | Where-Object name -eq "From").value
    $Subject = ($HeadersObj | Where-Object name -eq "Subject").value
    $Snippet = $EmailDetail.snippet

    # Rule-based categorization
    $Category = "Important" # Default
    $LowerSubject = $Subject.ToLower()
    $LowerSnippet = $Snippet.ToLower()


    if ($LowerSnippet -match "unsubscribe|newsletter|digest") {
        $Category = "Newsletter"
    } elseif ($LowerSubject -match "sale|offer|discount|promo|deal") {
        $Category = "Marketing"
    } elseif ($LowerSubject -match "opportunity|partnership|collaboration|invest|proposal") {
        $Category = "Biz Opp"
    }

    $Emails += [PSCustomObject]@{
        From = $From
        Subject = $Subject
        Snippet = $Snippet
        Category = $Category
    }
}

# Display summary
Write-Host "`n=== DAILY EMAIL SUMMARY ($($Emails.Count) new emails) ===" -ForegroundColor Cyan
$Emails | Format-Table Category, From, Subject -AutoSize


# Category counts
Write-Host "`nCategory Breakdown:" -ForegroundColor Yellow
$Emails | Group-Object Category | Select-Object Name, Count | Format-Table -AutoSize

# Optional: Save to file
$Emails | Export-Csv -Path "C:\SandboxShare\DailySummary_$(Get-Date -Format 'yyyMMdd').csv" -NoTypeInformation
Write-Host "Full report saved to CSV." -ForegroundColor Green