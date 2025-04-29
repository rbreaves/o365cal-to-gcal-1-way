$TokenPath = "$env:USERPROFILE\gmail-calendar-token.json"

# === LOAD CONFIG FROM .env ===
$envPath = ".\.env"
if (Test-Path $envPath) {
    Get-Content $envPath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]+?)\s*=\s*(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            Set-Variable -Name $key -Value $value -Scope Script
        }
    }
} else {
    Write-Host "❌ .env file not found! Exiting." -ForegroundColor Red
    exit 1
}

# === AUTH FUNCTION (REUSE FROM MAIN SCRIPT) ===
function Get-OAuth2Token {
    param (
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Scopes
    )

    $AuthUrl = "https://accounts.google.com/o/oauth2/v2/auth"
    $TokenUrl = "https://oauth2.googleapis.com/token"
    $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $authRequestUrl = "$AuthUrl?client_id=$ClientId&redirect_uri=$RedirectUri&response_type=code&scope=$Scopes&access_type=offline&prompt=consent"
    Write-Host "`nPlease log in to Google using the URL below:" -ForegroundColor Yellow
    Write-Host $authRequestUrl -ForegroundColor Cyan

    try { Start-Process "chrome.exe" $authRequestUrl }
    catch { try { Start-Process "msedge.exe" $authRequestUrl } catch { Start-Process "firefox.exe" $authRequestUrl } }

    $authCode = Read-Host "`nPaste the code you received after authorizing"

    $body = @{
        code = $authCode
        client_id = $ClientId
        client_secret = $ClientSecret
        redirect_uri = $RedirectUri
        grant_type = "authorization_code"
    }

    $response = Invoke-RestMethod -Method POST -Uri $TokenUrl -Body $body
    return $response
}

function Get-GmailAccessToken {
    if (Test-Path $TokenPath) {
        $TokenData = Get-Content $TokenPath | ConvertFrom-Json
    } else {
        $TokenData = Get-OAuth2Token `
            -ClientId $ClientID `
            -ClientSecret $ClientSecret `
            -Scopes "https://www.googleapis.com/auth/calendar"
        $TokenData | ConvertTo-Json | Out-File $TokenPath
    }
    return $TokenData.access_token
}

# === CONNECT TO GOOGLE CALENDAR ===
$AccessToken = Get-GmailAccessToken
$headers = @{Authorization = "Bearer $AccessToken"}

# Fetch all events (you can adjust the time range if you want)
$start = (Get-Date).AddYears(-1)
$end = (Get-Date).AddYears(1)
$calendarUrl = "https://www.googleapis.com/calendar/v3/calendars/$GmailCalendarID/events?timeMin=$($start.ToString("yyyy-MM-ddTHH:mm:ssZ"))&timeMax=$($end.ToString("yyyy-MM-ddTHH:mm:ssZ"))&singleEvents=true"

try {
    $googleEventsResponse = Invoke-RestMethod -Uri $calendarUrl -Headers $headers -Method GET
    $googleEvents = $googleEventsResponse.items
} catch {
    Write-Host "`n❌ Failed to fetch Google Calendar events." -ForegroundColor Red
    exit 1
}

Write-Host "Found $($googleEvents.Count) total events in Google Calendar." -ForegroundColor Green

# === DELETE TARGET EVENTS ===

$deletedCount = 0
foreach ($gEvent in $googleEvents) {
    try {
        $deleteUrl = "https://www.googleapis.com/calendar/v3/calendars/$GmailCalendarID/events/$($gEvent.id)"
        Invoke-RestMethod -Uri $deleteUrl -Method DELETE -Headers $headers
        Write-Host "🗑 Deleted: $($gEvent.summary) at $($gEvent.start.dateTime)" -ForegroundColor Red
        $deletedCount++
    } catch {
        Write-Host "❌ Failed to delete event: $($gEvent.summary)" -ForegroundColor Yellow
    }
}


Write-Host ""
Write-Host "=====================" -ForegroundColor Cyan
Write-Host "🧹 Cleanup Summary:" -ForegroundColor Cyan
Write-Host "🗑 Total events deleted: $deletedCount"
Write-Host "=====================" -ForegroundColor Cyan
