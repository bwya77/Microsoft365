$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret' 
$verbosepreference = 'continue'

Function Connect-MSGraphAPI {
    param (
        [system.string]$AppID,
        [system.string]$TenantID,
        [system.string]$AppSecret
    )
    begin {
        $URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $AppID
            Client_Secret = $AppSecret
        } 
    }
    Process {
        Write-Verbose "Connecting to the Graph API"
        $Response = Invoke-RestMethod -Uri $URI -Method POST -Body $ReqTokenBody
    }
    End {
        $Response
    }
}

Function Get-MSGraphRequest {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken
    )
    begin {
        [System.Array]$allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $Uri
        }
    }
    process {
        write-verbose "GET request at endpoint: $Uri"
        $data = Invoke-RestMethod @ReqTokenBody
        while ($data.'@odata.nextLink') {
            $allPages += $data.value
            $ReqTokenBody.Uri = $data.'@odata.nextLink'
            $Data = Invoke-RestMethod @ReqTokenBody
            # to avoid throttling, the loop will sleep for 3 seconds
            Start-Sleep -Seconds 3
        }
        $allPages += $data.value
    }
    end {
        Write-Verbose "Returning all results"
        $allPages
    }
}

Function Set-TeamsPresence {
    param (
        [Parameter(Position = 0, mandatory = $true)]
        [system.string]$AccessToken,
        [Parameter(Position = 1, mandatory = $true)]
        [ValidateSet('Available', 'DoNotDisturb', 'Away', 'Busy')]
        [system.string]$Presence,
        [Parameter(Position = 3, mandatory = $true)]
        [system.string]$UserID,
        [Parameter(Position = 4, mandatory = $true)]
        [system.int32]$ExpirationDuration,
        [Parameter(Position = 5, mandatory = $false)]
        [system.string]$AppID
    )
    begin {
        if ($Presence -eq "Available") {
            Write-Verbose "Presence will be set to Available"
            $availability = "Available"
            $activity     = "Available"
        }
        if ($Presence -eq "DoNotDisturb") {
            Write-Verbose "Presence will be set to DoNotDisturb"
            $availability = "DoNotDisturb"
            $activity     = "Presenting"
        }
        if ($Presence -eq "Away") {
            Write-Verbose "Presence will be set to Away"
            $availability = "Away"
            $activity     = "Away"
        }
        if ($Presence -eq "Busy") {
            Write-Verbose "Presence will be set to Busy"
            $availability = "Busy"
            $activity     = "InACall"
        }

        Write-Verbose "Building API Call"
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Post"
            Uri     = "https://graph.microsoft.com/v1.0/users/$UserID/presence/microsoft.graph.setPresence"
            Body = @{
                sessionId          = $AppID
                availability       = $availability
                activity           = $activity
                expirationDuration = "PT$ExpirationDuration`M"
            } | ConvertTo-Json
        }
        Write-Verbose $ReqTokenBody.body
    }
    process {
        Write-Verbose "Making API Call"
        Invoke-RestMethod @ReqTokenBody
    }
}

$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret

write-verbose "Getting all users"
$Users = (Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/users").value
foreach ($i in $Users) {
    Try {
        Write-Verbose "Getting events for $($i.displayName)"
        Write-Verbose "Getting the current time in UTC"
        $UTCTime = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        Write-Verbose "Getting the current time in UTC + 1 hour"
        $UTCTimePlusHour = (Get-Date).ToUniversalTime().AddHours(1).ToString("yyyy-MM-ddTHH:mm:ssZ")
        (Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/users/$($i.id)/calendar/events?`$filter=start/dateTime ge '$UTCTime' and subject eq 'DND' and start/dateTime le '$UTCTimePlusHour'").value | foreach-object {
            if ($_.id) {
                # Get the UTC offset of the time zone, timezone is based on the time zone set on the event
                Write-Verbose "Getting the UTC offset of the time zone of the event"
                $targettzoffset = [System.TimeZoneInfo]::FindSystemTimeZoneById("$($_.originalStartTimeZone)").BaseUtcOffset
                Write-Verbose "UTC Offset: $targettzoffset"
                Write-Verbose "Getting the current time in the time zone of the event"
                $Now = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "$($_.originalStartTimeZone)")
                Write-Verbose "Current time in the time zone of the event: $Now"
                Write-Verbose "Getting the time until the event starts"
                $timeUntil = New-TimeSpan -Start $Now -End $_.start.dateTime.AddHours($targettzoffset.hours)
                Write-Verbose "Time until the event starts: $timeUntil"
                Write-Verbose "Getting the duration of the event"
                $duration = New-TimeSpan  -start $_.start.dateTime -end $_.end.dateTime
                Write-Verbose "Duration of the event: $duration"

                $myObject = [PSCustomObject]@{
                    title        = $_.subject
                    start        = $_.start.dateTime.AddHours($targettzoffset.hours)
                    end          = $_.end.dateTime.AddHours($targettzoffset.hours)
                    duration     = $_.end.dateTime - $_.start.dateTime
                    minutesUntil = ([string]$timeUntil.TotalMinutes)
                    timeZone     = $_.originalStartTimeZone  
                }
                write-verbose "event details"
                Write-Verbose ($myObject | ConvertTo-Json -Depth 10)
                [int]$minUntil = [string]($myObject.minutesUntil).split(".")[0]
                Write-Verbose "Minutes until event: $minUntil"
                if ($minUntil -le 5) {
                    Write-Verbose "Event '$($_.subject)' starts in 5 minutes or less, setting Presence to DoNotDisturb"
                    if ($duration.totalMinutes -gt 240) {
                        $expirationDuration = 240
                        Write-Verbose "Event Duration is greater than 4 hours, setting to 4 hours"
                    }
                    else {
                        Write-Verbose "Event Duration is less than 4 hours, setting to $duration"
                        $expirationDuration = $duration.totalMinutes
                    }
                    Write-Verbose "Setting Presence to DoNotDisturb for $expirationDuration minutes"
                    Set-TeamsPresence -AccessToken $tokenResponse.access_token -Presence DoNotDisturb -UserID $i.id -ExpirationDuration $expirationDuration -AppID $AppID
                    Write-Verbose "Presence set to DoNotDisturb for $expirationDuration minutes"
                }
                else {
                    write-verbose "Event '$($_.subject)' starts in more than 5 minutes, skipping"
                }
            }
        }
    }
    Catch {
        Write-Verbose "$($i.displayName) does not have a valid mailbox, skipping"
        continue
    }
}
