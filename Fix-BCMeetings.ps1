#requires -Version 5.1
<#
.SYNOPSIS
    Finds, and optionally fixes, your Outlook meetings that show at the wrong time since
    British Columbia moved to permanent UTC-7 (1 November 2026).

.DESCRIPTION
    Meetings booked before the change in the old "Pacific Time (US & Canada)" zone
    (Windows ID "Pacific Standard Time") are stored by Exchange at the time they would have
    had under the old daylight-saving rules. In the winter months they now show one hour
    late (for example 10:30 shows as 11:30).

    This script works on YOUR OWN calendar only, using Microsoft Graph with your own
    sign-in (delegated permissions). It never reads anyone else's mailbox.

      -Mode Report (default)  Lists affected meetings. Changes nothing.
      -Mode Fix               Lists them, shows exactly what will change, asks you to type
                              YES, then moves the meetings YOU organize to the same clock
                              time in the new B.C. zone. Writes an undo file first.
      -Mode Undo              Puts meetings back exactly as they were, from an undo file.

    Safety rules:
      * Only meetings you organize are changed; others are listed with the organizer to ask.
      * Each meeting is re-read just before it is changed and skipped if it changed since
        the scan, or if this script (or the SMUS Time Zone Fixer app) already fixed it.
      * A recurring series is changed as a whole, and only when it has no changed or
        cancelled occurrences (those would be reset by Outlook); such series are listed
        for you to fix in Outlook.
      * All-day events and meetings in any other time zone are never touched.

.PARAMETER Mode
    Report (default), Fix or Undo.

.PARAMETER StartDate
    First day to check (default 2026-11-01, when B.C. stopped changing clocks).

.PARAMETER EndDate
    Last day to check (default 2028-03-13).

.PARAMETER NewTimeZone
    Windows time zone ID to move meetings into (default "British Columbia Standard Time",
    shown in Windows as "(UTC-07:00) Pacific Time (British Columbia)").

.PARAMETER OldTimeZone
    Windows time zone ID of the meetings that need fixing (default "Pacific Standard Time").

.PARAMETER CsvPath
    Also save the report as a CSV file.

.PARAMETER UndoFile
    Fix: where to write the undo file (default: Documents\BC-Meetings-Undo-<date>.json).
    Undo: the undo file to read.

.PARAMETER ClientId
    Application (client) ID of your organization's own Entra app registration, if your
    tenant doesn't allow the standard "Microsoft Graph Command Line Tools" app. See README.

.PARAMETER TenantId
    Your tenant ID or domain (for example contoso.com). Used with -ClientId.

.EXAMPLE
    .\Fix-BCMeetings.ps1
    Report only: lists the meetings that show at the wrong time.

.EXAMPLE
    .\Fix-BCMeetings.ps1 -Mode Fix
    Shows the changes, asks for YES, fixes the meetings you organize, writes an undo file.

.EXAMPLE
    .\Fix-BCMeetings.ps1 -Mode Undo -UndoFile "$HOME\Documents\BC-Meetings-Undo-20261102-0930.json"

.LINK
    README.md in the same folder.
#>
[CmdletBinding()]
param(
    [ValidateSet('Report', 'Fix', 'Undo')]
    [string]$Mode = 'Report',
    [datetime]$StartDate = '2026-11-01',
    [datetime]$EndDate = '2028-03-13',
    [string]$NewTimeZone = 'British Columbia Standard Time',
    [string]$OldTimeZone = 'Pacific Standard Time',
    [string]$CsvPath,
    [string]$UndoFile,
    [string]$ClientId,
    [string]$TenantId
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

# The same extended properties as the SMUS Time Zone Fixer app, so the two recognise each
# other's fixes and never shift a meeting twice.
$PropSet = '{6f1c2b6e-3d4a-4c1e-9a52-8b7d2f0e4a11}'
$PropOriginal = "String $PropSet Name TzFixOriginal"
$PropState = "String $PropSet Name TzFixState"
$NewOffsetMinutes = -420   # B.C. is UTC-7 all year

# ---- pure helpers (tested by Test-FixBCMeetings.ps1) -----------------------------------------

function Get-OldPacificZone {
    <# The old Pacific rules (with daylight saving), on Windows or on macOS/Linux. #>
    foreach ($id in 'Pacific Standard Time', 'America/Los_Angeles') {
        try { return [TimeZoneInfo]::FindSystemTimeZoneById($id) } catch { }
    }
    throw 'Could not find the Pacific time zone on this computer.'
}

function ConvertFrom-GraphTime {
    <# Graph dateTime text such as "2026-11-10T16:30:00.0000000" -> DateTime (kind Utc). #>
    param([string]$Text)
    if (-not $Text -or $Text.Length -lt 19) { return $null }
    $value = [datetime]::ParseExact($Text.Substring(0, 19), 'yyyy-MM-ddTHH:mm:ss',
        [Globalization.CultureInfo]::InvariantCulture)
    return [datetime]::SpecifyKind($value, [DateTimeKind]::Utc)
}

function Format-Wall {
    <# DateTime -> "2026-11-10T08:30:00" (the clock time Graph expects with a timeZone). #>
    param([datetime]$Value)
    return $Value.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-Day {
    <# DateTime -> "2026-11-01T00:00:00Z" (midnight UTC of that date) for calendarView. #>
    param([datetime]$Value)
    return $Value.ToString('yyyy-MM-dd"T00:00:00Z"', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-MeetingStatus {
    <#
      Decide what to do with one calendar item (times read in UTC).
      Returns Fix | AskOrganizer | Check | OK, the clock time it shows now and the clock time
      it was booked for.
    #>
    param(
        [Parameter(Mandatory)] $Item,
        [Parameter(Mandatory)] [TimeZoneInfo]$OldZone,
        [Parameter(Mandatory)] [string]$OldZoneId,
        [Parameter(Mandatory)] [string]$NewZoneId,
        [string[]]$FixedIds = @()
    )
    $start = ConvertFrom-GraphTime $Item.start.dateTime
    $zone = [string]$Item.originalStartTimeZone
    $result = [ordered]@{ Status = 'OK'; ShowsNow = $null; BookedFor = $null; Reason = '' }
    if ($null -eq $start) { $result.Status = 'Check'; $result.Reason = 'No start time.'; return $result }
    $result.ShowsNow = $start.AddMinutes($NewOffsetMinutes)
    $result.BookedFor = [TimeZoneInfo]::ConvertTimeFromUtc($start, $OldZone)
    $fixedBefore = ($FixedIds -contains [string]$Item.id) -or
                   ($Item.seriesMasterId -and $FixedIds -contains [string]$Item.seriesMasterId)

    if ($Item.isAllDay -eq $true) { return $result }
    if ($zone -eq $NewZoneId) { return $result }
    if ($zone -ne $OldZoneId) {
        $result.Status = 'Check'
        $result.Reason = "Saved in another time zone ($(if ($zone) { $zone } else { 'none' }))."
        return $result
    }
    if ($result.ShowsNow -eq $result.BookedFor) { return $result }   # summer dates: still right
    if ($fixedBefore) {
        $result.Status = 'Check'
        $result.Reason = 'Fixed before, then changed in Outlook. Check the time.'
        return $result
    }
    if ($Item.isOrganizer -eq $true) { $result.Status = 'Fix' }
    else { $result.Status = 'AskOrganizer'; $result.Reason = 'Only the organizer can change it.' }
    return $result
}

function Get-StatusRank([string]$Status) {
    switch ($Status) { 'Fix' { 1 } 'AskOrganizer' { 2 } 'Check' { 3 } default { 4 } }
}

# ---- Microsoft Graph -------------------------------------------------------------------------

function Invoke-Graph {
    param([string]$Method = 'GET', [string]$Uri, $Body, [string]$TimeZoneHeader = 'UTC')
    $headers = @{ Prefer = "outlook.timezone=`"$TimeZoneHeader`"" }
    $params = @{ Method = $Method; Uri = $Uri; Headers = $headers; OutputType = 'Hashtable' }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress)
        $params.ContentType = 'application/json'
    }
    return Invoke-MgGraphRequest @params
}

function Get-GraphPages {
    param([string]$Uri, [string]$TimeZoneHeader = 'UTC')
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Uri $next -TimeZoneHeader $TimeZoneHeader
        foreach ($item in $page.value) { $items.Add($item) }
        $next = $page.'@odata.nextLink'
    }
    return , $items
}

function Get-EventPath([string]$Id) { "v1.0/me/events/$([uri]::EscapeDataString($Id))" }

function Get-Prop($Item, [string]$PropId) {
    foreach ($p in @($Item.singleValueExtendedProperties)) {
        if ($p -and $p.id -and $p.id.ToLower() -eq $PropId.ToLower()) { return [string]$p.value }
    }
    return $null
}

function Connect-Calendar {
    param([string[]]$Scopes)
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw ("The Microsoft Graph PowerShell module isn't installed. Run:`n" +
               "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser")
    }
    $connect = @{ Scopes = $Scopes; NoWelcome = $true }
    if ($ClientId) { $connect.ClientId = $ClientId }
    if ($TenantId) { $connect.TenantId = $TenantId }
    Connect-MgGraph @connect
    $context = Get-MgContext
    Write-Host "Signed in as $($context.Account)" -ForegroundColor Green
    try { [void](Invoke-Graph -Uri 'v1.0/me/calendar?$select=id') }
    catch {
        throw ("Signed in as $($context.Account), but that account has no mailbox Graph can " +
               "read (MailboxNotEnabledForRESTAPI). Sign in with your own work account: " +
               "run Disconnect-MgGraph and try again.")
    }
}

function Get-FixedIds {
    <# Ids of meetings already fixed by this script or by the SMUS Time Zone Fixer app. #>
    $filter = "singleValueExtendedProperties/any(ep: ep/id eq '$PropState' and ep/value eq 'fixed')"
    try {
        $items = Get-GraphPages "v1.0/me/events?`$filter=$([uri]::EscapeDataString($filter))&`$select=id"
        return @($items | ForEach-Object { [string]$_.id })
    } catch {
        Write-Warning "Couldn't look up meetings fixed before; continuing without."
        return @()
    }
}

function Get-Meetings {
    $zone = Get-OldPacificZone
    $s = Format-Day $StartDate; $e = Format-Day $EndDate
    $select = 'subject,start,end,isAllDay,type,seriesMasterId,originalStartTimeZone,organizer,isOrganizer,attendees,isCancelled'
    Write-Host "Reading your calendar from $($StartDate.ToString('d MMM yyyy')) to $($EndDate.ToString('d MMM yyyy'))..."
    $events = Get-GraphPages "v1.0/me/calendarView?startDateTime=$s&endDateTime=$e&`$select=$select&`$top=100"
    $fixed = Get-FixedIds

    # One row per meeting: a recurring series counts once (its first wrong occurrence shown).
    $groups = @{}
    foreach ($ev in $events) {
        if ($ev.isCancelled -eq $true) { continue }
        $st = Get-MeetingStatus -Item $ev -OldZone $zone -OldZoneId $OldTimeZone -NewZoneId $NewTimeZone -FixedIds $fixed
        $key = if ($ev.seriesMasterId) { [string]$ev.seriesMasterId } else { [string]$ev.id }
        $row = [pscustomobject]@{
            Key        = $key
            Id         = [string]$ev.id
            Recurring  = [bool]$ev.seriesMasterId
            Type       = [string]$ev.type
            Subject    = if ($ev.subject) { [string]$ev.subject } else { '(no subject)' }
            Organizer  = [string]$ev.organizer.emailAddress.address
            Attendees  = @($ev.attendees).Count
            Status     = $st.Status
            ShowsNow   = $st.ShowsNow
            BookedFor  = $st.BookedFor
            Reason     = $st.Reason
        }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($row)
    }
    $meetings = foreach ($key in $groups.Keys) {
        $rows = $groups[$key]
        $best = $rows | Sort-Object { Get-StatusRank $_.Status }, ShowsNow | Select-Object -First 1
        $best | Add-Member -NotePropertyName HasChangedWeeks -NotePropertyValue ([bool]($rows | Where-Object Type -eq 'exception')) -PassThru
    }
    return @($meetings | Sort-Object { Get-StatusRank $_.Status }, ShowsNow)
}

function Show-Report($Meetings) {
    $fix = @($Meetings | Where-Object Status -eq 'Fix')
    $ask = @($Meetings | Where-Object Status -eq 'AskOrganizer')
    $check = @($Meetings | Where-Object Status -eq 'Check')
    $ok = @($Meetings | Where-Object Status -eq 'OK')
    $fmt = { param($m) '{0,-40} {1,-24} {2:ddd d MMM, h:mm tt} -> {3:h:mm tt}' -f
             ($(if ($m.Subject.Length -gt 40) { $m.Subject.Substring(0, 37) + '...' } else { $m.Subject })),
             ($(if ($m.Recurring) { 'recurring' } else { '' })), $m.ShowsNow, $m.BookedFor }

    Write-Host ''
    Write-Host "You can fix: $($fix.Count)" -ForegroundColor Yellow
    foreach ($m in $fix) { Write-Host ('  ' + (& $fmt $m)) }
    if ($ask.Count) {
        Write-Host "Ask the organizer: $($ask.Count)" -ForegroundColor Yellow
        foreach ($m in $ask) { Write-Host ('  ' + (& $fmt $m) + "   organizer: $($m.Organizer)") }
    }
    if ($check.Count) {
        Write-Host "Check yourself in Outlook: $($check.Count)" -ForegroundColor Yellow
        foreach ($m in $check) { Write-Host ('  ' + (& $fmt $m) + "   $($m.Reason)") }
    }
    Write-Host "Already correct: $($ok.Count)" -ForegroundColor Green
    if ($CsvPath) {
        $Meetings | Select-Object Status, Subject, Recurring, Organizer, Attendees,
            @{ n = 'ShowsNow'; e = { $_.ShowsNow.ToString('yyyy-MM-dd HH:mm') } },
            @{ n = 'BookedFor'; e = { $_.BookedFor.ToString('yyyy-MM-dd HH:mm') } }, Reason |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report saved to $CsvPath"
    }
}

# ---- fixing ----------------------------------------------------------------------------------

function Get-SeriesProblem {
    <# Why a series can't be fixed safely as a whole, or $null. #>
    param($Master, [string]$MasterId)
    if (-not $Master.recurrence) { return 'no repeat pattern found' }
    if (@($Master.cancelledOccurrences).Count -gt 0) { return 'has cancelled occurrences' }
    $from = [datetime]::Parse($Master.recurrence.range.startDate, [Globalization.CultureInfo]::InvariantCulture)
    $to = $EndDate
    while ($from -lt $to) {
        $chunkEnd = $from.AddDays(90); if ($chunkEnd -gt $to) { $chunkEnd = $to }
        $uri = "$(Get-EventPath $MasterId)/instances?startDateTime=$(Format-Day $from)" +
               "&endDateTime=$(Format-Day $chunkEnd)&`$select=type&`$top=200"
        if (Get-GraphPages $uri | Where-Object { $_.type -eq 'exception' }) { return 'has changed occurrences' }
        $from = $chunkEnd
    }
    return $null
}

function Invoke-FixOne($Meeting) {
    <# Returns a result object; adds an undo record to $script:UndoRecords before changing. #>
    $select = 'subject,start,end,isOrganizer,originalStartTimeZone,recurrence,type,cancelledOccurrences'
    $expand = "singleValueExtendedProperties(`$filter=id eq '$PropState')"
    $live = Invoke-Graph -Uri "$(Get-EventPath $Meeting.Key)?`$select=$select&`$expand=$expand" -TimeZoneHeader $OldTimeZone
    if ((Get-Prop $live $PropState) -eq 'fixed') { return 'skipped: already fixed' }
    if ($live.isOrganizer -ne $true) { return 'skipped: you are not the organizer' }
    if ([string]$live.originalStartTimeZone -ne $OldTimeZone) { return 'skipped: changed since the scan' }

    $isSeries = [string]$live.type -eq 'seriesMaster'
    if (-not $isSeries) {
        # Read in the old zone, the clock time IS the booked time; it must still match the scan.
        if ($live.start.dateTime.Substring(0, 19) -ne (Format-Wall $Meeting.BookedFor)) {
            return 'skipped: changed since the scan'
        }
    } else {
        $problem = Get-SeriesProblem -Master $live -MasterId $Meeting.Key
        if ($problem) { return "skipped: recurring meeting $problem - fix it in Outlook" }
    }

    $start = $live.start.dateTime.Substring(0, 19); $end = $live.end.dateTime.Substring(0, 19)
    $original = [ordered]@{
        kind  = $(if ($isSeries) { 'series' } else { 'single' })
        start = @{ dateTime = $start; timeZone = $OldTimeZone }
        end   = @{ dateTime = $end; timeZone = $OldTimeZone }
    }
    $body = @{
        start = @{ dateTime = $start; timeZone = $NewTimeZone }
        end   = @{ dateTime = $end; timeZone = $NewTimeZone }
    }
    if ($isSeries) {
        # The undo record keeps its own copy; the live pattern gets the new zone.
        $original.recurrence = $live.recurrence | ConvertTo-Json -Depth 10 | ConvertFrom-Json
        $live.recurrence.range.recurrenceTimeZone = $NewTimeZone
        $body.recurrence = $live.recurrence
    }
    $body.singleValueExtendedProperties = @(
        @{ id = $PropOriginal; value = ($original | ConvertTo-Json -Depth 10 -Compress) },
        @{ id = $PropState; value = 'fixed' })

    $script:UndoRecords.Add([ordered]@{ id = $Meeting.Key; subject = $Meeting.Subject; original = $original;
                                        fixedStart = $start; newTimeZone = $NewTimeZone })
    Save-UndoFile
    Invoke-Graph -Method PATCH -Uri (Get-EventPath $Meeting.Key) -Body $body | Out-Null
    return 'fixed'
}

function Save-UndoFile {
    $script:UndoRecords | ConvertTo-Json -Depth 20 | Set-Content -Path $script:UndoPath -Encoding UTF8
}

function Invoke-Fix($Meetings) {
    $fix = @($Meetings | Where-Object Status -eq 'Fix')
    if (-not $fix.Count) { Write-Host "`nNothing for you to fix." -ForegroundColor Green; return }
    $people = ($fix | Measure-Object Attendees -Sum).Sum
    Write-Host ''
    Write-Host "This will move $($fix.Count) meeting(s) to the same clock time in '$NewTimeZone'." -ForegroundColor Cyan
    Write-Host "Recurring meetings are changed as a whole series. About $people attendee(s) get an updated invitation."
    Write-Host "An undo file is written to: $script:UndoPath"
    $answer = Read-Host 'Type YES to fix them, anything else to stop'
    if ($answer -cne 'YES') { Write-Host 'Stopped. Nothing was changed.'; return }

    $results = foreach ($m in $fix) {
        try { $outcome = Invoke-FixOne $m }
        catch { $outcome = "failed: $($_.Exception.Message)" }
        $color = if ($outcome -eq 'fixed') { 'Green' } else { 'Yellow' }
        Write-Host ("  {0,-40} {1}" -f $m.Subject, $outcome) -ForegroundColor $color
        [pscustomobject]@{ Subject = $m.Subject; Result = $outcome }
    }
    $done = @($results | Where-Object Result -eq 'fixed').Count
    Write-Host "`nFixed $done of $($fix.Count). Undo file: $script:UndoPath" -ForegroundColor Cyan
    Write-Host 'Open Outlook and check a few of them. Run the script again: they should show as already correct.'
}

function Invoke-Undo {
    if (-not $UndoFile -or -not (Test-Path $UndoFile)) { throw 'Give the undo file with -UndoFile <path>.' }
    $records = @(Get-Content -Raw $UndoFile | ConvertFrom-Json)
    Write-Host "Undo file has $($records.Count) meeting(s)."
    $answer = Read-Host 'Type YES to put them back as they were'
    if ($answer -cne 'YES') { Write-Host 'Stopped. Nothing was changed.'; return }
    foreach ($r in $records) {
        try {
            $expand = "singleValueExtendedProperties(`$filter=id eq '$PropState')"
            $live = Invoke-Graph -Uri "$(Get-EventPath $r.id)?`$select=start,originalStartTimeZone,recurrence&`$expand=$expand" -TimeZoneHeader $r.newTimeZone
            if ((Get-Prop $live $PropState) -ne 'fixed' -or [string]$live.originalStartTimeZone -ne $r.newTimeZone -or
                $live.start.dateTime.Substring(0, 19) -ne $r.fixedStart) {
                Write-Host ("  {0,-40} skipped: changed since the fix" -f $r.subject) -ForegroundColor Yellow
                continue
            }
            $body = @{ start = $r.original.start; end = $r.original.end;
                       singleValueExtendedProperties = @(@{ id = $PropState; value = 'undone' }) }
            if ($r.original.kind -eq 'series') { $body.recurrence = $r.original.recurrence }
            Invoke-Graph -Method PATCH -Uri (Get-EventPath $r.id) -Body $body | Out-Null
            Write-Host ("  {0,-40} undone" -f $r.subject) -ForegroundColor Green
        } catch {
            Write-Host ("  {0,-40} failed: {1}" -f $r.subject, $_.Exception.Message) -ForegroundColor Red
        }
    }
}

# ---- main ------------------------------------------------------------------------------------

# Dot-sourcing (". .\Fix-BCMeetings.ps1") only loads the functions, for the tests.
if ($MyInvocation.InvocationName -eq '.') { return }

Write-Host "Fix-BCMeetings $ScriptVersion  ($Mode mode)" -ForegroundColor Cyan
if ($Mode -eq 'Report') {
    # With your own app registration, ask for what it was granted (see README).
    Connect-Calendar -Scopes $(if ($ClientId) { 'Calendars.ReadWrite' } else { 'Calendars.Read' })
    Show-Report (Get-Meetings)
    Write-Host "`nNothing was changed. To fix the meetings you organize, run with -Mode Fix."
} elseif ($Mode -eq 'Fix') {
    if (-not $UndoFile) {
        $UndoFile = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("BC-Meetings-Undo-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmm'))
    }
    $script:UndoPath = $UndoFile
    $script:UndoRecords = New-Object System.Collections.Generic.List[object]
    Connect-Calendar -Scopes 'Calendars.ReadWrite'
    $meetings = Get-Meetings
    Show-Report $meetings
    Invoke-Fix $meetings
} else {
    Connect-Calendar -Scopes 'Calendars.ReadWrite'
    Invoke-Undo
}
