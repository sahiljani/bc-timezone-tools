#requires -Version 5.1
<#
.SYNOPSIS
    Offline self-test of Fix-BCMeetings.ps1's time logic. No sign-in, no calendar access.

.EXAMPLE
    .\Test-FixBCMeetings.ps1
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Fix-BCMeetings.ps1')   # loads the functions only

$script:failed = 0
function Assert-Equal($Actual, $Expected, [string]$Name) {
    if ($Actual -eq $Expected) { Write-Host "  PASS  $Name" -ForegroundColor Green }
    else { Write-Host "  FAIL  $Name (expected '$Expected', got '$Actual')" -ForegroundColor Red; $script:failed++ }
}

function New-TestItem([string]$Utc, [string]$Zone = 'Pacific Standard Time', [bool]$Organizer = $true,
                  [bool]$AllDay = $false, [string]$Id = 'A1', [string]$Master = $null) {
    @{ id = $Id; seriesMasterId = $Master; isAllDay = $AllDay; isOrganizer = $Organizer
       originalStartTimeZone = $Zone; start = @{ dateTime = "$Utc.0000000"; timeZone = 'UTC' } }
}

$zone = Get-OldPacificZone
$common = @{ OldZone = $zone; OldZoneId = 'Pacific Standard Time'; NewZoneId = 'British Columbia Standard Time' }

Write-Host 'Fix-BCMeetings self-test'

# Booked 8:30 in November under the old rules = 16:30 UTC; B.C. now shows it at 9:30.
$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00') @common
Assert-Equal $r.Status 'Fix' 'winter meeting I organize needs fixing'
Assert-Equal $r.BookedFor.ToString('HH:mm') '08:30' 'booked for 8:30'
Assert-Equal $r.ShowsNow.ToString('HH:mm') '09:30' 'shows at 9:30'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00' -Organizer $false) @common
Assert-Equal $r.Status 'AskOrganizer' 'someone else''s meeting: ask the organizer'

$r = Get-MeetingStatus -Item (New-TestItem '2027-07-12T16:30:00') @common
Assert-Equal $r.Status 'OK' 'summer date: old and new rules agree'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00' -Zone 'British Columbia Standard Time') @common
Assert-Equal $r.Status 'OK' 'already in the B.C. zone'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00' -Zone 'Eastern Standard Time') @common
Assert-Equal $r.Status 'Check' 'another time zone: check yourself'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T00:00:00' -AllDay $true) @common
Assert-Equal $r.Status 'OK' 'all-day events are never touched'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00' -Id 'X' -Master 'M1') -FixedIds @('M1') @common
Assert-Equal $r.Status 'Check' 'fixed before then changed back: never shifted twice'

$r = Get-MeetingStatus -Item (New-TestItem '2026-11-10T16:30:00' -Id 'Y') -FixedIds @('M1') @common
Assert-Equal $r.Status 'Fix' 'a single meeting is not confused with a fixed series'

Assert-Equal (Format-Day ([datetime]'2026-11-01')) '2026-11-01T00:00:00Z' 'calendar dates in UTC'
Assert-Equal (ConvertFrom-GraphTime '2026-11-10T16:30:00.0000000').Kind 'Utc' 'Graph times read as UTC'

if ($script:failed) { Write-Host "$script:failed test(s) failed." -ForegroundColor Red; exit 1 }
Write-Host 'All tests passed.' -ForegroundColor Green
