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


# === AUTH FUNCTIONS ===
function Get-OAuth2Token {
    param (
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Scopes
    )

    $AuthUrl = "https://accounts.google.com/o/oauth2/v2/auth"
    $TokenUrl = "https://oauth2.googleapis.com/token"
    $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $authRequestUrl = "$AuthUrl`?client_id=$ClientId`&redirect_uri=$RedirectUri&response_type=code&scope=$Scopes&access_type=offline&prompt=consent"

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

function Refresh-AccessToken {
    param (
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RefreshToken
    )

    $TokenUrl = "https://oauth2.googleapis.com/token"
    $body = @{
        client_id = $ClientId
        client_secret = $ClientSecret
        refresh_token = $RefreshToken
        grant_type = "refresh_token"
    }

    $response = Invoke-RestMethod -Method POST -Uri $TokenUrl -Body $body
    return $response
}

function Get-GmailAccessToken {
    if (Test-Path $TokenPath) {
        try {
            $TokenData = Get-Content -Path $TokenPath -Raw | ConvertFrom-Json
        } catch {
            Write-Host "`n⚠️ Failed to read or parse token file. Reauthenticating..." -ForegroundColor Yellow
            $TokenData = Get-OAuth2Token -ClientId $ClientID -ClientSecret $ClientSecret -Scopes "https://www.googleapis.com/auth/calendar"
            $TokenData | ConvertTo-Json -Depth 3 | Set-Content -Encoding UTF8 -Path $TokenPath
        }
    } else {
        $TokenData = Get-OAuth2Token -ClientId $ClientID -ClientSecret $ClientSecret -Scopes "https://www.googleapis.com/auth/calendar"
        $TokenData | ConvertTo-Json -Depth 3 | Set-Content -Encoding UTF8 -Path $TokenPath
    }

    # Sanity check
    if (-not $TokenData.access_token) {
        Write-Host "`n❌ No access token present. Exiting." -ForegroundColor Red
        exit 1
    }

    # Test the current access token
    $headers = @{ Authorization = "Bearer $($TokenData.access_token)" }
    try {
        Invoke-RestMethod -Uri "https://www.googleapis.com/calendar/v3/calendars/primary" -Headers $headers -Method GET -ErrorAction Stop
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Write-Host "Access token expired - refreshing..." -ForegroundColor Yellow
            if ($TokenData.refresh_token) {
                try {
                    $NewTokenData = Invoke-RestMethod -Method POST -Uri "https://oauth2.googleapis.com/token" -Body @{
                        client_id     = $ClientID
                        client_secret = $ClientSecret
                        refresh_token = $TokenData.refresh_token
                        grant_type    = "refresh_token"
                    }

                    $TokenData.access_token = $NewTokenData.access_token
                    $TokenData.expires_in = $NewTokenData.expires_in
                    $TokenData.token_type = $NewTokenData.token_type

                    $TokenData | ConvertTo-Json -Depth 3 | Set-Content -Encoding UTF8 -Path $TokenPath
                } catch {
                    Write-Host "Failed to refresh access token." -ForegroundColor Red
                    exit 1
                }
            } else {
                Write-Host "No refresh token available. Please delete the token file and reauthenticate." -ForegroundColor Red
                exit 1
            }
        } else {
            Write-Host "Unexpected error when validating token: $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    return $TokenData.access_token
}


# === OUTLOOK SETUP ===
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$start = (Get-Date).AddDays(-30)
$end = (Get-Date).AddDays(30)

# Personal calendar
$personalCalendar = $namespace.GetDefaultFolder(9)

# Shared calendar: Loaded from .env
try {
    $eaMailbox = $namespace.Folders.Item($SharedCalendarMailboxName)
    $sharedCalendar = $eaMailbox.Folders.Item("Calendar")
} catch {
    $sharedCalendar = $null
}


# === FUNCTION TO FETCH OUTLOOK EVENTS ===
function Get-CalendarEvents($calendarFolder, $start, $end) {
    $items = $calendarFolder.Items
    $items.IncludeRecurrences = $true
    $items.Sort("[Start]")
    $filter = "[Start] >= '" + $start.ToString("g") + "' AND [End] <= '" + $end.ToString("g") + "'"
    $restrictedItems = $items.Restrict($filter)

    $events = @()
    foreach ($item in $restrictedItems) {
        try {
            $eventInfo = [PSCustomObject]@{
                Subject = $item.Subject
                Start = (Get-Date $item.Start -Format "yyyy-MM-ddTHH:mm:ss")
                End = (Get-Date $item.End -Format "yyyy-MM-ddTHH:mm:ss")
            }
            $events += $eventInfo
        } catch {}
    }
    return $events
}

# === FETCH OUTLOOK EVENTS ===
$personalEvents = Get-CalendarEvents -calendarFolder $personalCalendar -start $start -end $end
$sharedEvents = @()
if ($sharedCalendar) {
    $sharedEvents = Get-CalendarEvents -calendarFolder $sharedCalendar -start $start -end $end
}

$allEvents = $personalEvents + $sharedEvents
Write-Host "Outlook Events to sync: $($allEvents.Count)" -ForegroundColor Green

# === FETCH GOOGLE CALENDAR EVENTS ===
$AccessToken = Get-GmailAccessToken
$headers = @{Authorization = "Bearer $AccessToken"}

$calendarUrl = "https://www.googleapis.com/calendar/v3/calendars/$GmailCalendarID/events?timeMin=$($start.ToString("yyyy-MM-ddTHH:mm:ssZ"))&timeMax=$($end.ToString("yyyy-MM-ddTHH:mm:ssZ"))&singleEvents=true"

try {
    $googleEventsResponse = Invoke-RestMethod -Uri $calendarUrl -Headers $headers -Method GET
    $googleEvents = $googleEventsResponse.items
} catch {
    Write-Host "`n❌ Failed to fetch Google Calendar events." -ForegroundColor Red
    if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
        Write-Host "Unauthorized (401) — your access token is invalid or expired." -ForegroundColor Yellow
        Write-Host "Try deleting your token file and re-running the script to reauthorize:" -ForegroundColor Yellow
        Write-Host "    Remove-Item '$TokenPath'`n" -ForegroundColor Yellow
    } elseif ($_.Exception.Response.StatusCode.Value__ -eq 403) {
        Write-Host "Forbidden (403) — your calendar may not be shared with this app or token lacks permission." -ForegroundColor Yellow
    } else {
        Write-Host "Unhandled error: $($_.Exception.Message)" -ForegroundColor Red
    }
    exit 1
}

if (-not $googleEvents) {
    $googleEvents = @()
}

Write-Host "Google events in calendar: $($googleEvents.Count)" -ForegroundColor Green

# === SYNC EVENTS ===
$addedCount = 0
$skippedCount = 0
$deletedCount = 0

# Build existing Google Event keys based on custom description
$googleKeys = @{}
foreach ($gEvent in $googleEvents) {
    if ($gEvent.PSObject.Properties["description"] -and $gEvent.description -like "SyncedFromOutlook*") {
        $googleKeys[$gEvent.description] = $gEvent.id
    }
}

# === ADD NEW EVENTS TO GMAIL ===
foreach ($event in $allEvents) {
    $eventKey = "SyncedFromOutlook|$($event.Subject)|$($event.Start)"

    if (-not $googleKeys.ContainsKey($eventKey)) {
        $startTime = Get-Date $event.Start -Format "yyyy-MM-ddTHH:mm:ss"
        $endTime = Get-Date $event.End -Format "yyyy-MM-ddTHH:mm:ss"

        $eventBody = @{
            summary = $event.Subject
            description = $eventKey
            start = @{
                dateTime = $startTime
                timeZone = $GoogleCalendarTimeZone
            }
            end = @{
                dateTime = $endTime
                timeZone = $GoogleCalendarTimeZone
            }
        } | ConvertTo-Json -Depth 3 -Compress

        Invoke-RestMethod -Method POST -Uri "https://www.googleapis.com/calendar/v3/calendars/$GmailCalendarID/events" `
            -Headers $headers -Body $eventBody -ContentType "application/json"

        Write-Host "➕ Added: $($event.Subject) at $($startTime)" -ForegroundColor Yellow
        $addedCount++
    } else {
        $skippedCount++
    }
}


# === DELETE ORPHANED GMAIL EVENTS ===
$outlookKeys = @{}
foreach ($event in $allEvents) {
    $outlookKeys["SyncedFromOutlook|$($event.Subject)|$($event.Start)"] = $true
}

foreach ($key in $googleKeys.Keys) {
    if (-not $outlookKeys.ContainsKey($key)) {
        $deleteUrl = "https://www.googleapis.com/calendar/v3/calendars/$GmailCalendarID/events/$($googleKeys[$key])"
        try {
            Invoke-RestMethod -Uri $deleteUrl -Method DELETE -Headers $headers
            Write-Host "🗑 Deleted orphan event with key: $key" -ForegroundColor Red
            $deletedCount++
        } catch {}
    }
}

# === SUMMARY ===
Write-Host ""
Write-Host "=====================" -ForegroundColor Cyan
Write-Host "📋 Sync Summary:" -ForegroundColor Cyan
Write-Host "✅ Added: $addedCount"
Write-Host "🔁 Skipped (already present): $skippedCount"
Write-Host "🗑 Deleted orphans: $deletedCount"
Write-Host "=====================" -ForegroundColor Cyan