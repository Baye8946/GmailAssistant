# Test-GmailRead.ps1 - List recent emails using refresh token

$ClientJson = Get-Content -Path "C:\SandboxShare\client_secret.json" -Raw | ConvertFrom-Json
$ClientID = $ClientJson.installed.client_id
$ClientSecret = $ClientJson.installed.client_secret

$TokenJson = Get-Content -Path "C:\SandboxShare\gmail_token.json" -Raw | ConvertFrom-Json
$RefreshToken = $TokenJson.refresh_token

# Refresh access token
$RefreshBody = @{
    client_id = $ClientID
    client_secret = $ClientSecret
    refresh_token = $RefreshToken
    grant_type = "refresh_token"
}

$RefreshResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $RefreshBody -ContentType "application/x-www-form-urlencoded"

$AccessToken = $RefreshResponse.access_token

# List recent emails
$Headers = @{
    Authorization = "Bearer $AccessToken"
}

$Messages = Invoke-RestMethod -Uri "https://www.googleapis.com/gmail/v1/users/me/messages?maxResults=10" -Headers $Headers -Method Get

Write-Host "Recent 10 Email IDs:" -ForegroundColor Green
$Messages.messages.id | ForEach-Object { Write-Host $_ }

# Optional: Get first email details
if ($Messages.messages) {
    $FirstId = $Messages.messages[0].id
    $Email = Invoke-RestMethod -Uri "https://gmail.googleapis.com/gmail/v1/users/me/messages/$FirstId" -Headers $Headers -Method Get
    $Subject = ($Email.payload.headers | Where-Object { $_.name -eq "From" }).value
    Write-Host "Latest Email: From $From - Subject: $Subject" -ForegroundColor Cyan
}
