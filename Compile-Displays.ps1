<#
EHS Room Displays - nightly compiler.

Reads the registry workbook and every teacher's userid.xlsx from the synced
SharePoint library, validates them, writes one data.json next to the site,
and publishes it. Screens pick it up automatically - no site edits, ever.

ONE-TIME SETUP (on the always-on PC, logged in as the account that will run it):
  1. Sync the SharePoint library in OneDrive; right-click the library folder ->
     "Always keep on this device" (files-on-demand placeholders break scripts).
  2. Install-Module ImportExcel -Scope CurrentUser      (reads xlsx WITHOUT Excel;
     safe for scheduled/headless runs)
  3. git clone your ehs-display repo to $SitePath and do one manual push so the
     credential manager stores your GitHub token.
  4. Edit the CONFIG block below.
  5. Schedule it (run as YOUR user, "whether user is logged on or not"):
       schtasks /Create /TN "EHS Displays Nightly" /SC DAILY /ST 04:15 ^
         /TR "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ehs-display\Compile-Displays.ps1"
  Same-day publish: double-click "Publish Now.bat" any time.

The report (issues + what was compiled) is written into the library so it is
visible in SharePoint. A failed course keeps yesterday's data on the wall
instead of going blank. A fatal error (registry unreadable) publishes nothing
and leaves the previous data.json untouched.
#>

# ============================ CONFIG ============================
$LibraryPath   = "$env:USERPROFILE\Edwardsville CUSD 7\CTE Displays - Documents"  # synced SharePoint library
$RegistryFile  = 'EHS_room_display_data.xlsx'   # office-owned: Settings, Bell, DayTypes, Rooms
$TeacherFolder = 'teachers'                     # subfolder of userid.xlsx files
$SitePath      = 'C:\ehs-display'               # local clone of the site repo (data.json lands here)
$PublishMode   = 'github'                       # 'github' | 'copy' | 'none'
$WebRootPath   = '\\district-server\web\displays'  # used when PublishMode = 'copy'
$ReportName    = 'compile-report.txt'
$AnnouncementsUrl = 'https://ehs.ecusd7.org/link-2'   # school Daily Announcements page; blank = skip the scrape
$GitExe        = 'git'                          # or a full path, e.g. 'C:\ehs-display\PortableGit\cmd\git.exe' if Git couldn't be installed
# ================================================================

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $LibraryPath)) {
    Write-Host ""
    Write-Host "STOP: the library folder was not found:" -ForegroundColor Red
    Write-Host "      $LibraryPath"
    Write-Host "Fix `$LibraryPath in the CONFIG block at the top of this script. It must be the synced"
    Write-Host "'CTE Displays' folder as it appears in File Explorer (open it, click the address bar, copy)."
    Write-Host "If no such folder exists yet, sync the channel folder from SharePoint first (Sync button)."
    exit 1
}
if (-not (Test-Path (Join-Path $LibraryPath $RegistryFile))) {
    Write-Host ""
    Write-Host "STOP: the registry workbook is not in the library folder:" -ForegroundColor Red
    Write-Host "      $(Join-Path $LibraryPath $RegistryFile)"
    Write-Host "Upload EHS_room_display_data.xlsx to the CTE Displays folder in the channel and let OneDrive sync it."
    exit 1
}
Import-Module ImportExcel

$issues  = New-Object System.Collections.Generic.List[string]
$noteOk  = New-Object System.Collections.Generic.List[string]

function Normalize-Key([string]$h) {
    (([string]$h) -replace '[^A-Za-z0-9]+', '_').Trim('_').ToLower()
}
function Format-Cell($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [datetime]) {
        if ($v.TimeOfDay -eq [timespan]::Zero) { return $v.ToString('yyyy-MM-dd') }
        return $v.ToString('h:mm tt')
    }
    return ([string]$v).Trim()
}
function Read-Tab([string]$path, [string]$sheet) {
    # Registry and teacher tabs both have a note in row 1 and headers in row 2.
    $rows = Import-Excel -Path $path -WorksheetName $sheet -StartRow 2
    $out = @()
    foreach ($r in $rows) {
        $o = [ordered]@{}; $any = $false
        foreach ($p in $r.PSObject.Properties) {
            $k = Normalize-Key $p.Name
            if (-not $k) { continue }
            $v = Format-Cell $p.Value
            $o[$k] = $v
            if ($v) { $any = $true }
        }
        if ($any) { $out += [pscustomobject]$o }
    }
    return ,$out
}

function Get-MonthStart([string]$tf, [int]$mo, $firstDay) {
    # first Monday of the month; "mid" -> second Monday, "late" -> third Monday
    $yr = if ($mo -ge 7) { $firstDay.Year } else { $firstDay.Year + 1 }
    $d = Get-Date -Year $yr -Month $mo -Day 1
    while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    if ($tf -match '(?i)\bmid') { $d = $d.AddDays(7) } elseif ($tf -match '(?i)\blate') { $d = $d.AddDays(14) }
    if ($d -lt $firstDay) { $d = $firstDay }
    return $d.ToString('yyyy-MM-dd')
}
function Read-SnsCourse([string]$path, [hashtable]$weekStarts, $firstDay, [hashtable]$anchors) {
    # District scope & sequence template: 9 unit blocks of 12 rows.
    # A = Time Frame, B = Chapter/Unit (first line becomes the wall title), C = Learning Target(s).
    # C-cell lines starting "I can" (or "PF: I can") become success criteria; the rest are targets.
    # If a cell is ALL I-cans, they are promoted to the targets (that IS the unit's wall content).
    $pkg = Open-ExcelPackage -Path $path
    try {
        $ws = $pkg.Workbook.Worksheets[1]     # one course per S&S file: always the first sheet
        $units = @()
        for ($k = 0; $k -lt 9; $k++) {
            $r = 3 + 12 * $k
            $unitCell = [string]$ws.Cells["B$r"].Text
            if (-not $unitCell.Trim()) { continue }
            $unit = ($unitCell -split "`n")[0].Trim()
            $tf = ([string]$ws.Cells["A$r"].Text) -replace "`n", ' '
            $noteCell = [string]$ws.Cells["H$r"].Text
            $lt = @(); $sc = @()
            foreach ($l in (([string]$ws.Cells["C$r"].Text) -split "`n")) {
                $l = $l.Trim(); if (-not $l) { continue }
                if ($l -match '^(PF:\s*)?I can') { $sc += $l } else { $lt += $l }
            }
            if (-not $lt.Count) { $lt = $sc; $sc = @() }
            $starts = ''
            if ($tf -match '[Ww]eeks?\s*(\d+)') {
                $wk = [int]$Matches[1]
                if ($weekStarts.ContainsKey($wk)) { $starts = $weekStarts[$wk] }
                elseif ($firstDay) { $starts = $firstDay.AddDays(7 * ($wk - 1)).ToString('yyyy-MM-dd') }
                # "Weeks 17-20 (Jan)": if the teacher also wrote a month and the week lands elsewhere,
                # the month is their intent - departments count weeks differently than the calendar does.
                if ($starts -and $firstDay -and $tf -match '(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b') {
                    $mo = @{ jan=1; feb=2; mar=3; apr=4; may=5; jun=6; jul=7; aug=8; sep=9; oct=10; nov=11; dec=12 }[$Matches[1].ToLower()]
                    if (([datetime]::Parse($starts)).Month -ne $mo) { $starts = Get-MonthStart $tf $mo $firstDay }
                }
            } elseif ($tf -match '(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})') {
                $y = [int]$Matches[3]; if ($y -lt 100) { $y += 2000 }
                $starts = (Get-Date -Year $y -Month ([int]$Matches[1]) -Day ([int]$Matches[2])).ToString('yyyy-MM-dd')
            } elseif ($tf -match '(?i)\b(q|quarter\s*)([1-4])\b' -and $anchors -and $anchors.ContainsKey('q' + $Matches[2])) {
                $starts = $anchors['q' + $Matches[2]]
            } elseif ($tf -match '(?i)\b(s|sem(ester)?\s*)([12])\b' -and $anchors -and $anchors.ContainsKey('s' + $Matches[3])) {
                $starts = $anchors['s' + $Matches[3]]
            } elseif ($firstDay -and $tf -match '(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b') {
                # a month name alone ("Aug", "late Oct", "Nov-Dec"): the first Monday of that month in the school year
                $mo = @{ jan=1; feb=2; mar=3; apr=4; may=5; jun=6; jul=7; aug=8; sep=9; oct=10; nov=11; dec=12 }[$Matches[1].ToLower()]
                $starts = Get-MonthStart $tf $mo $firstDay
            }
            $cur = ''
            if ($noteCell -match '^(?i)\s*(x|now|current)\b') { $cur = 'x' }   # pacing override from the S&S Notes cell
            $units += [pscustomobject]@{ unit = $unit; starts = $starts; learning_target_s = ($lt -join "`n"); success_criteria = ($sc -join "`n"); current = $cur; notes = '' }
        }
        # keep units in sequence: if two resolve to the same week, the later one starts a week after the earlier
        $prevStart = $null
        foreach ($u in $units) {
            if (-not $u.starts) { continue }
            $d = [datetime]::Parse($u.starts)
            if ($prevStart -and $d -le $prevStart) { $d = $prevStart.AddDays(7); $u.starts = $d.ToString('yyyy-MM-dd') }
            $prevStart = $d
        }
        return ,$units
    } finally { Close-ExcelPackage $pkg -NoSave }
}

function Scrape-Announcements([string]$url) {
    # The school's Daily Announcements page: a date header "ANNOUNCEMENTS <date>" in bold,
    # then one paragraph per item, each led by a bold phrase. Returns items as
    # @{ date; title; text }. Throws on fetch failure; returns @() if nothing parses.
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }   # Windows PowerShell 5.1 on older builds
    $html = (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) EHS-Displays-Compiler').Content
    $start = $html.IndexOf('fsPageContent'); if ($start -lt 0) { $start = 0 }
    $body = $html.Substring($start)
    $stop = $body.IndexOf('fsFooter'); if ($stop -gt 0) { $body = $body.Substring(0, $stop) }
    $date = ''; $items = @()
    foreach ($m in [regex]::Matches($body, '<p[^>]*>(.*?)</p>', 'Singleline')) {
        $inner = $m.Groups[1].Value
        $lm = [regex]::Match($inner, '<(?:strong|b)[^>]*>(.*?)</(?:strong|b)>', 'Singleline')
        $leadTxt = ''
        if ($lm.Success) {
            $leadTxt = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($lm.Groups[1].Value, '<[^>]+>', ' '))
            $leadTxt = ([regex]::Replace($leadTxt, '\s+', ' ')).Trim()
        }
        $flat = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($inner, '<[^>]+>', ' '))
        $flat = ([regex]::Replace($flat, '\s+', ' ')).Trim()
        if (-not $flat) { continue }
        if ($leadTxt -match '^ANNOUNCEMENTS\s+(.+)$') { $date = $Matches[1].Trim(); continue }
        if (-not $leadTxt -or $leadTxt.Length -gt 140 -or $flat.Length -gt 1500) { continue }   # nav/footer junk
        $rest = $flat
        if ($flat.StartsWith($leadTxt)) { $rest = $flat.Substring($leadTxt.Length).TrimStart(' ', ',', ':', ';', '-', '.') }
        $items += @{ date = $date; title = $leadTxt; text = $rest }
    }
    return ,$items
}

function Normalize-Course([string]$t) {
    # "AP Macroeconomics (LCCC ECON 151) - Full Year" -> "apmacroeconomics"
    $t = [string]$t
    $t = $t -replace '\([^)]*\)', ' '                      # drop parentheticals
    $t = ($t -split '\s+[-\u2013\u2014]\s+')[0]           # drop " - suffix" / " – suffix" / " — suffix"
    return ($t -replace '[^A-Za-z0-9]', '').ToLower()
}
function Test-Current([string]$v) { return ([string]$v).Trim() -match '^(?i)(x|yes|y|true|1|now|current)$' }

function Build-SnsIndex([string]$folder) {
    # Filenames do not matter. Every .xlsx in sns\ is opened and asked which course it is:
    # the district template carries the course name in row 1 (next to the "Course Name:" label).
    $map = @{}; $notes = @()
    if (-not (Test-Path $folder)) { return @{ map = $map; notes = $notes } }
    foreach ($f in Get-ChildItem -Path $folder -Filter *.xlsx -File -Recurse) {
        if ($f.Name -like '~$*') { continue }
        $name = ''
        try {
            $pkg = Open-ExcelPackage -Path $f.FullName
            try {
                $ws = $pkg.Workbook.Worksheets[1]
                foreach ($addr in @('C1','D1','B1','E1','C2','B2')) {
                    $t = ([string]$ws.Cells[$addr].Text).Trim()
                    if ($t -and $t -notmatch '^(?i)(course\s*name|last\s*update|time\s*frame)') { $name = $t; break }
                }
            } finally { Close-ExcelPackage $pkg -NoSave }
        } catch {
            $notes += "sns\$($f.Name): could not be opened ($($_.Exception.Message))."
            continue
        }
        if (-not $name) { $notes += "sns\$($f.Name): no course name found in row 1 - put the course name next to 'Course Name:'."; continue }
        $k = Normalize-Course $name
        if ($map.ContainsKey($k)) { $notes += "sns\$($f.Name): duplicate course '$name' (already provided by $(Split-Path $map[$k].file -Leaf)) - using the first."; continue }
        $map[$k] = @{ file = $f.FullName; label = $name }
    }
    return @{ map = $map; notes = $notes }
}

# ---------- previous compile (per-course safety net) ----------
$dataPath = Join-Path $SitePath 'data.json'
$prev = $null
if (Test-Path $dataPath) {
    try { $prev = Get-Content $dataPath -Raw | ConvertFrom-Json } catch { $prev = $null }
}

# ---------- registry (fatal if unreadable) ----------
$reg = Join-Path $LibraryPath $RegistryFile
try {
    $settingsRows = Read-Tab $reg 'Settings'
    $bell         = Read-Tab $reg 'Bell'
    $daytypes     = Read-Tab $reg 'DayTypes'
    $rooms        = Read-Tab $reg 'Rooms'
} catch {
    $msg = "FATAL: cannot read registry '$reg' - $($_.Exception.Message). Nothing published; previous data.json untouched."
    $msg | Set-Content (Join-Path $LibraryPath $ReportName)
    Write-Host $msg
    exit 1
}
$settings = [ordered]@{}
foreach ($r in $settingsRows) { if ($r.key) { $settings[$r.key] = $r.value } }
$weekStarts = @{}    # school week number -> Monday date (registry Weeks tab; survives breaks)
try { foreach ($w in (Read-Tab $reg 'Weeks')) { if ($w.week -and $w.starts) { $weekStarts[[int]$w.week] = ([datetime]::Parse($w.starts)).ToString('yyyy-MM-dd') } } } catch { }
$anchors = @{ q1 = '2026-08-19'; q2 = '2026-10-26'; q3 = '2027-01-19'; q4 = '2027-03-22' }
foreach ($k in @('q1','q2','q3','q4')) { if ($settings[$k + '_start']) { $anchors[$k] = ([datetime]::Parse($settings[$k + '_start'])).ToString('yyyy-MM-dd') } }
$anchors['s1'] = $anchors['q1']; $anchors['s2'] = $anchors['q3']
$sns = Build-SnsIndex (Join-Path $LibraryPath 'sns')
foreach ($n in $sns.notes) { $issues.Add($n) }
$snsUsed = @{}
$firstDay = $null
if ($settings['first_day']) { try { $firstDay = [datetime]::Parse($settings['first_day']) } catch { $issues.Add("Settings: first_day '$($settings['first_day'])' is unreadable.") } }

# ---------- teacher course tabs ----------
$courses = [ordered]@{}
# Teachers tab: teacher number -> userid (typed once, survives every semester's Rooms regeneration)
$uidByNumber = @{}
try { foreach ($t in (Read-Tab $reg 'Teachers')) { if ($t.teacher_number -and $t.userid) { $uidByNumber[[string]$t.teacher_number] = $t.userid } } } catch { }
foreach ($r in $rooms) {
    if (-not $r.teacher_userid -and $r.teacher_number -and $uidByNumber.ContainsKey([string]$r.teacher_number)) { $r | Add-Member -NotePropertyName teacher_userid -NotePropertyValue $uidByNumber[[string]$r.teacher_number] -Force }
}
$needs = @{}; $srcMap = @{}   # course -> userid / optional scope & sequence source (first non-blank wins)
foreach ($r in $rooms) {
    if ($r.course) {
        if (-not $needs.ContainsKey($r.course)) { $needs[$r.course] = ''; $srcMap[$r.course] = '' }
        if ($r.teacher_userid -and -not $needs[$r.course]) { $needs[$r.course] = $r.teacher_userid }
        if ($r.content_source -and -not $srcMap[$r.course]) { $srcMap[$r.course] = $r.content_source }
    }
}
$today = (Get-Date).Date
$noSource = @(); $missingByUid = @{}
foreach ($course in ($needs.Keys | Sort-Object)) {
    $uid = $needs[$course]; $src = $srcMap[$course]
    $rowsOut = $null
    $snsKey = Normalize-Course $course
    if (-not $src) {
        $hit = $null
        if ($sns.map.ContainsKey($snsKey)) { $hit = $snsKey }
        else {   # tolerant: the file's name starts with the registry name, or vice versa ("AP Macro" vs "AP Macroeconomics")
            foreach ($k in $sns.map.Keys) { if ($k.StartsWith($snsKey) -or $snsKey.StartsWith($k)) { $hit = $k; break } }
        }
        if ($hit) { $src = $sns.map[$hit].file; $snsUsed[$hit] = $true }   # matched by the course name INSIDE the file
    }
    if ($src) {
        $file = if ([System.IO.Path]::IsPathRooted($src)) { $src } else { Join-Path $LibraryPath $src }
        if (-not (Test-Path $file)) {
            $issues.Add("Missing scope & sequence file: $src (course '$course').")
        } else {
            try { $rowsOut = Read-SnsCourse $file $weekStarts $firstDay $anchors }
            catch { $issues.Add("S&S file '$src' unreadable ($($_.Exception.Message))."); $rowsOut = $null }
        }
    } elseif (-not $uid) {
        $noSource += $course
    } else {
        $file = Join-Path (Join-Path $LibraryPath $TeacherFolder) ($uid + '.xlsx')
        if (-not (Test-Path $file)) {
            if (-not $missingByUid.ContainsKey($uid)) { $missingByUid[$uid] = @() }
            $missingByUid[$uid] += $course
        } else {
            $tabs = @()
            try { $tabs = @((Get-ExcelSheetInfo -Path $file).Name) } catch { }
            if ($tabs.Count -and ($tabs -notcontains $course)) {
                $issues.Add("$uid.xlsx: no tab named '$course' - run Provision-Teachers.ps1 to add it (tabs found: $($tabs -join ', ')).")
            } else {
                try { $rowsOut = Read-Tab $file $course }
                catch { $issues.Add("$uid.xlsx: tab '$course' unreadable ($($_.Exception.Message))."); $rowsOut = $null }
            }
        }
    }
    if ($null -ne $rowsOut) {
        $srcName = if ($src) { $src } else { "$uid.xlsx [$course]" }
        $active = $false; $xUnits = @(); $calUnit = ''; $calDate = $null
        foreach ($u in $rowsOut) {
            if (-not $u.unit) { continue }
            if (Test-Current $u.current) { $active = $true; $xUnits += $u.unit }
            $d = $null
            if ($u.starts) {
                try { $d = [datetime]::Parse($u.starts) } catch { $issues.Add("${srcName}: unit '$($u.unit)' has an unreadable Starts date '$($u.starts)'.") }
            } elseif (-not (Test-Current $u.current)) {
                $issues.Add("${srcName}: unit '$($u.unit)' can't be scheduled - add its week to the registry Weeks tab, use a date, or mark Current x.")
            }
            if ($d -and $d -le $today) { $active = $true; if (-not $calDate -or $d -gt $calDate) { $calDate = $d; $calUnit = $u.unit } }
            if (((Test-Current $u.current) -or ($d -and $d -le $today)) -and -not $u.learning_target_s) {
                $issues.Add("${srcName}: unit '$($u.unit)' is live but has no learning targets.")
            }
        }
        if ($xUnits.Count -gt 1) { $issues.Add("${srcName}: $($xUnits.Count) units are marked current - the wall shows the last one, '$($xUnits[-1])'. Clear the extras.") }
        if ($xUnits.Count -ge 1 -and $calUnit -and $xUnits[-1] -ne $calUnit) { $noteOk.Add("${srcName}: pacing override active - showing '$($xUnits[-1])' while the calendar says '$calUnit'.") }
        if (-not $active) { $issues.Add("${srcName}: no unit is active today (check dates/weeks or mark one Current x).") }
    }
    if ($null -eq $rowsOut) {
        if ($prev -and $prev.courses.$course) {
            $courses[$course] = $prev.courses.$course
            $noteOk.Add("'$course': kept yesterday's data on the wall.")
        } else {
            $courses[$course] = @()
        }
    } else {
        $courses[$course] = $rowsOut
    }
}

if ($noSource.Count) {
    $shown = ($noSource | Select-Object -First 8) -join '; '
    if ($noSource.Count -gt 8) { $shown += "; ... $($noSource.Count - 8) more" }
    $noteOk.Add("$($noSource.Count) course(s) have no scope & sequence file and no teacher userid yet - their walls show the announcements deck. ($shown)")
}
foreach ($u in ($missingByUid.Keys | Sort-Object)) {
    $issues.Add("Missing file: teachers\$u.xlsx - run Provision-Teachers.ps1 (courses: $($missingByUid[$u] -join '; ')).")
}
foreach ($k in $sns.map.Keys) {
    if (-not $snsUsed.ContainsKey($k)) {
        $issues.Add("sns\$(Split-Path $sns.map[$k].file -Leaf): course name '$($sns.map[$k].label)' does not match any Course in the registry - fix the name in the file or the Rooms tab.")
    }
}

# ---------- daily announcements (scraped from the school site) ----------
$announcements = @()
if ($AnnouncementsUrl) {
    try {
        $announcements = @(Scrape-Announcements $AnnouncementsUrl)
        if ($announcements.Count) { $noteOk.Add("Announcements: $($announcements.Count) items from the school site ($($announcements[0].date)).") }
        else { throw 'page fetched but no items parsed - layout may have changed' }
    } catch {
        $issues.Add("Announcements scrape failed ($($_.Exception.Message)) - keeping yesterday's.")
        if ($prev -and $prev.announcements) { $announcements = @($prev.announcements) }
    }
}

# ---------- slim the PUBLISHED copy ----------
# data.json is served from a PUBLIC web page. The registry keeps every column;
# only what a screen actually renders is published. Dropped here:
#   notes            section numbers, enrolled/cap, "shares period"  (debug only)
#   teacher_userid   a staff login name - never publish
#   source_sheet_id / content_source   internal file wiring
# Teacher names stay, in friendly form (Mr. Koester), because the hub lists and
# searches by teacher. DEBUGGING: if a screen shows the wrong course, the full
# detail is in the registry's Rooms tab and in compile-report.txt - not here.
function Get-Honorific([string]$teacher) {
    # "Koester, A." -> "Mr. Koester" is not knowable from the SIS, so we publish
    # the surname with the initial and no title: "A. Koester".
    $t = ([string]$teacher).Trim()
    if ($t -match '^\s*([^,]+),\s*(\S)') { return "$($Matches[2]). $($Matches[1])" }
    return $t
}
$publicRooms = @()
foreach ($r in $rooms) {
    $publicRooms += [pscustomobject]@{
        room         = $r.room
        period       = $r.period
        course       = $r.course
        course_title = $r.course_title
        teacher      = Get-Honorific $r.teacher
        day_type     = $r.day_type
    }
}
$noteOk.Add("Published data.json slimmed: notes/enrollment, teacher userids and file paths stay in the registry only.")

# ---------- assemble + write ----------
$out = [ordered]@{
    generated = (Get-Date).ToString('s')
    settings  = $settings
    bell      = $bell
    daytypes  = $daytypes
    rooms     = $publicRooms
    courses   = $courses
    announcements = $announcements
}
if (-not (Test-Path $SitePath)) { New-Item -ItemType Directory -Path $SitePath | Out-Null }
$out | ConvertTo-Json -Depth 8 | Set-Content -Path $dataPath -Encoding UTF8
# version.txt: screens poll this tiny file and only download data.json when it changes.
Set-Content -Path (Join-Path $SitePath 'version.txt') -Value $out.generated -Encoding ASCII -NoNewline

# ---------- publish ----------
switch ($PublishMode) {
    'github' {
        & $GitExe -C $SitePath add data.json version.txt | Out-Null
        & $GitExe -C $SitePath diff --cached --quiet
        if ($LASTEXITCODE -ne 0) {
            & $GitExe -C $SitePath commit -m ("data " + (Get-Date -Format 'yyyy-MM-dd HH:mm')) | Out-Null
            & $GitExe -C $SitePath push | Out-Null
            $noteOk.Add('Published to GitHub Pages.')
        } else { $noteOk.Add('No changes since last compile - nothing pushed.') }
    }
    'copy' {
        Copy-Item $dataPath (Join-Path $WebRootPath 'data.json') -Force
        Copy-Item (Join-Path $SitePath 'version.txt') (Join-Path $WebRootPath 'version.txt') -Force
        $noteOk.Add("Copied data.json to $WebRootPath.")
    }
    default { $noteOk.Add('PublishMode none - data.json written locally only.') }
}

# ---------- report (visible in SharePoint) ----------
$report = @()
$report += "EHS Displays compile - $(Get-Date)"
$report += "Rooms: $($rooms.Count) rows | Courses compiled: $($courses.Keys.Count) | Issues: $($issues.Count)"
$report += ''
if ($issues.Count) { $report += 'ISSUES (fix in the named file; the wall shows last good data meanwhile):'; $report += $issues } else { $report += 'No issues. Every course compiled clean.' }
$report += ''
$report += $noteOk
$report -join "`r`n" | Set-Content (Join-Path $LibraryPath $ReportName)
# Teachers never see compile-report.txt, so the sns-related lines also land where they look:
$snsIssues = @($issues | Where-Object { $_ -like 'sns\*' })
$snsNote = Join-Path (Join-Path $LibraryPath 'sns') '_READ ME - problems found last night.txt'
if ($snsIssues.Count) {
    (@("Checked $(Get-Date -Format 'ddd MMM d, h:mm tt'). Each line names the file and what to fix; your wall shows the last good version until then.", '') + $snsIssues) -join "`r`n" | Set-Content $snsNote
} elseif (Test-Path $snsNote) { Remove-Item $snsNote -Force }
Write-Host ($report -join "`n")
