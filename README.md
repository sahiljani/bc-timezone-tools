# Fix-BCMeetings

A free PowerShell script that finds, and optionally fixes, Outlook / Microsoft 365 meetings
that show at the **wrong time** since **British Columbia moved to permanent UTC‑7 on
1 November 2026**.

- Works on **your own calendar only**, with your own sign-in.
- **Report mode changes nothing.** Fixing only happens when you ask for it and type `YES`.
- Writes an **undo file** before changing anything.
- Runs on Windows PowerShell 5.1 and PowerShell 7 (Windows, macOS, Linux).

---

## The problem

Before 1 November 2026, B.C. used "Pacific Time (US & Canada)" (Windows ID
`Pacific Standard Time`), which springs forward and falls back. Meetings booked in that zone
are stored at the time they would have had under the old daylight‑saving rules.

B.C. now stays on UTC‑7 all year (Windows: **"(UTC‑07:00) Pacific Time (British Columbia)"**,
ID `British Columbia Standard Time`). So between November and mid‑March, those meetings show
**one hour late**:

| Booked as | Shows now (Nov–Mar) | Shows now (Mar–Nov) |
|---|---|---|
| 10:30 Pacific Time (US & Canada) | **11:30** ❌ | 10:30 ✅ |

The fix is to move each meeting to the **same clock time in the new B.C. zone**. That is what
this script does: it does not "shift by one hour", so summer dates and repeat patterns stay
right.

---

## Quick start

1. **Install the Microsoft Graph sign-in module** (once):
   ```powershell
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   ```
   Answer **Y** if it asks about an untrusted repository (PSGallery).

2. **Download** `Fix-BCMeetings.ps1` to a folder, open PowerShell in that folder, and unblock it:
   ```powershell
   Unblock-File .\Fix-BCMeetings.ps1
   ```

3. **See which meetings are affected** (changes nothing):
   ```powershell
   .\Fix-BCMeetings.ps1
   ```
   A Microsoft sign-in window opens: sign in with **your own work account**.

4. **Fix the meetings you organize** (shows the list, asks you to type `YES`):
   ```powershell
   .\Fix-BCMeetings.ps1 -Mode Fix
   ```

5. Open Outlook and check a few meetings. Run step 3 again: they now show as already correct.

> If PowerShell says *"running scripts is disabled on this system"*, run this once and try again:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

---

## What you'll see

```
Fix-BCMeetings 1.0.0  (Report mode)
Signed in as jane.doe@example.org
Reading your calendar from 1 Nov 2026 to 13 Mar 2028...

You can fix: 2
  Staff meeting                                                     Tue 10 Nov, 11:30 AM -> 10:30 AM
  Grade 9 team                             recurring                Mon 9 Nov, 9:00 AM -> 8:00 AM
Ask the organizer: 1
  Board meeting                                                     Wed 11 Nov, 4:00 PM -> 3:00 PM   organizer: chair@example.org
Already correct: 14

Nothing was changed. To fix the meetings you organize, run with -Mode Fix.
```

`Shows now -> Booked for`: the first time is what Outlook shows today, the second is what the
meeting was booked for and will show after fixing.

---

## Modes and options

| Command | What it does |
|---|---|
| `.\Fix-BCMeetings.ps1` | **Report** (default). Lists affected meetings. Changes nothing. |
| `.\Fix-BCMeetings.ps1 -CsvPath report.csv` | Report, and also saves it as a CSV file. |
| `.\Fix-BCMeetings.ps1 -Mode Fix` | Report, then asks you to type `YES`, then fixes the meetings **you organize**. |
| `.\Fix-BCMeetings.ps1 -Mode Undo -UndoFile <file>` | Puts fixed meetings back exactly as they were. |

| Option | Default | Meaning |
|---|---|---|
| `-StartDate` | `2026-11-01` | First day to check. |
| `-EndDate` | `2028-03-13` | Last day to check. |
| `-NewTimeZone` | `British Columbia Standard Time` | Zone to move meetings into. |
| `-OldTimeZone` | `Pacific Standard Time` | Zone of the meetings that need fixing. |
| `-UndoFile` | `Documents\BC-Meetings-Undo-<date>.json` | Fix: where to write the undo file. Undo: which file to read. |
| `-ClientId`, `-TenantId` | — | Use your organization's own app registration (see *Sign-in and permissions*). |

---

## How it decides

Each meeting between the start and end dates is put in one group:

| Group | When | What the script does |
|---|---|---|
| **You can fix** | Saved in the old Pacific zone, shows the wrong time, and **you are the organizer** | Fix mode moves it to the same clock time in the B.C. zone. |
| **Ask the organizer** | Same, but someone else organized it | Lists the organizer's email. Only they can change it for everyone. |
| **Check yourself** | Saved in another zone, or fixed before and then changed again | Never changed. Open it in Outlook and check the time. |
| **Already correct** | Already in the B.C. zone, all-day, or a summer date where old and new rules agree | Left alone. |

A **recurring meeting counts once**. It is fixed as a **whole series in one step** (one update
email to attendees), and only if it has **no changed or cancelled occurrences**: Outlook would
reset those when the series changes. Such a series is reported as *skipped* so you can fix it
in Outlook (open the series, set the time zone to "Pacific Time (British Columbia)", keep the
time).

---

## Safety

- **Your calendar only.** The script uses Microsoft Graph `/me` with delegated permissions.
  It cannot read or change anyone else's mailbox.
- **Report mode is read-only** (`Calendars.Read`). Fix and Undo ask for `Calendars.ReadWrite`.
- **Nothing changes without `YES`.** You see the full list first.
- **Re-checked before every change.** Each meeting is read again just before it is changed and
  skipped if it changed since the scan, if you're no longer the organizer, or if it was
  already fixed.
- **Never shifts a meeting twice.** Fixed meetings get a hidden marker
  (`TzFixState = fixed`). The marker is shared with the *SMUS Time Zone Fixer* desktop app, so
  the two tools recognise each other's fixes.
- **Undo file first.** Before each change, the original start, end, zone and repeat pattern are
  written to the undo file (and to a hidden property on the meeting).
- **Attendees** get the normal Outlook "meeting updated" email, once per meeting or series.

### Undo

```powershell
.\Fix-BCMeetings.ps1 -Mode Undo -UndoFile "$HOME\Documents\BC-Meetings-Undo-20261102-0930.json"
```

Each meeting is only put back if it is still exactly as the fix left it; anything changed
since is skipped and reported.

The undo file contains meeting subjects and times. Keep it private, and delete it when you no
longer need it.

---

## Sign-in and permissions

By default the script signs in through Microsoft's own **"Microsoft Graph Command Line Tools"**
app. Many organizations allow this; in others you'll see **"Need admin approval"**. Then either:

### Option A: an admin approves the Microsoft app

A Global or Application Administrator signs in once and consents for the organization:

```powershell
Connect-MgGraph -Scopes Calendars.ReadWrite
```

…and ticks **"Consent on behalf of your organization"**. Each person still only reaches their
own calendar.

### Option B: use your organization's own app registration (recommended for IT)

1. **Microsoft Entra admin center** → **Applications** → **App registrations** → **New registration**.
   - Name: for example `BC Meetings Fixer`
   - Supported account types: **Accounts in this organizational directory only**
   - Register.
2. **Authentication** → **Add a platform** → **Mobile and desktop applications** → tick
   `http://localhost` → **Configure**. Set **Allow public client flows** to **Yes** → **Save**.
3. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** →
   `Calendars.ReadWrite` → **Add**. Then **Grant admin consent for <your organization>**.
4. From **Overview**, copy the **Application (client) ID** and **Directory (tenant) ID**.
5. Run the script with them:
   ```powershell
   .\Fix-BCMeetings.ps1 -ClientId <client-id> -TenantId <tenant-id>
   .\Fix-BCMeetings.ps1 -ClientId <client-id> -TenantId <tenant-id> -Mode Fix
   ```

With your own app, Report mode also requests `Calendars.ReadWrite` (the permission the app was
granted), but it still only reads.

---

## Troubleshooting

| Message | What to do |
|---|---|
| *running scripts is disabled on this system* | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then `Unblock-File .\Fix-BCMeetings.ps1`. |
| *The Microsoft Graph PowerShell module isn't installed* | `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser` |
| *Need admin approval* | See *Sign-in and permissions* above. |
| *…has no mailbox Graph can read (MailboxNotEnabledForRESTAPI)* | You signed in with an admin or other account without a mailbox. Run `Disconnect-MgGraph`, then run the script again and choose your own work account. |
| *skipped: recurring meeting has changed occurrences* | Fix that series in Outlook: open the series, set the time zone to "Pacific Time (British Columbia)", keep the time. |
| *skipped: changed since the scan* | The meeting was edited while the script ran. Run it again. |
| *The time zone 'British Columbia Standard Time' is invalid* (PATCH failed) | Your Microsoft 365 tenant doesn't know the new zone yet. Wait for Microsoft's update and try again later. |
| Wrong account signs in automatically | `Disconnect-MgGraph`, then run again. Windows may offer the last account; choose *Use another account*. |

---

## Test the script (offline)

`Test-FixBCMeetings.ps1` checks the time logic without signing in or touching any calendar:

```powershell
.\Test-FixBCMeetings.ps1
```

---

## Limitations

- Only meetings saved in `Pacific Standard Time` are fixed. Meetings saved in other zones are
  reported for you to check.
- Recurring meetings with changed or cancelled occurrences are not fixed automatically.
- Meetings you didn't organize can only be fixed by their organizer.
- Only dates between `-StartDate` and `-EndDate` are checked.
- The script changes calendar items, not your computer's or Outlook's time zone setting. Set
  Windows to **"(UTC‑07:00) Pacific Time (British Columbia)"** (Settings → Time & language →
  Date & time) and Outlook to the same zone.

---

## License

MIT. See [LICENSE](LICENSE). Provided "as is", without warranty: run Report mode first and
check the results before fixing.
