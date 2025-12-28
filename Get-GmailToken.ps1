# Get-GmailToken.ps1 - First-time OAuth authorization for Gmail API
# Run this once to get refresh token


$ClientJson = Get-Content -Path "C:\SandboxShare\client_secret.json" -Raw | ConvertFrom-Json
$ClientID = $ClientJson.installed.client_id
$ClientSecret = $ClientJson.installed.client_secret

# Scopes for full functionality
$Scopes = @(
    "https://www.googleapis.com/auth/gmail.readonly"
    "https://www.googleapis.com/auth/gmail.modify"
    "https://www.googleapis.com/auth/gmail.send"
    "https://www.googleapis.com/auth/gmail.labels"
)
$ScopeEncoded = ($Scopes -join " ") | ForEach-Object { [Uri]::EscapeDataString($_) } | ForEach-Object { $_ -replace "\+", "%20" }

$RedirectUri = "urn:ietf:wg:oauth:2.0:oob" # For desktop apps (copy-paste code)

$AuthUrl = "https://accounts.google.com/o/oauth2/v2/auth?" +
           "client_id=$ClientID" +
           "&redirect_uri=$RedirectUri" +
           "&response_type=code" +
           "&scope=$ScopeEncoded" +
           "&access_type=offline" +
           "&prompt=consent" # Forces refresh token

Write-Host "Opening browser for authorization..." -ForegroundColor Green
Start-Process $AuthUrl

$Code = Read-Host "Paste the authorization code from the browser here"

$TokenBody = @{
    code = $Code
    client_id = $ClientID
    client_secret = $ClientSecret
    redirect_uri = $RedirectUri
    grant_type = "authorization_code"
}

$TokenResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $TokenBody -ContentType "application/x-www-form-urlencoded"

$TokenResponse | ConvertTo-Json | Out-File -FilePath "C:\SandboxShare\gmail_token.json" -Encoding utf8

Write-Host "Tokens saved to C:\SandboxShare\gmail_token.json" -ForegroundColor Green
Write-Host "IMPORTANT: This file contains your refresh_token - keep it secure!" -ForegroundColor Yellow
Write-Host "First access_token (expires in 1 hour): $($TokenResponse.access_token)"