param(
  [int]$KeepDays = 35,
  [int]$MotorWindowDays = 30,
  [int]$BackfillDays = 30,
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataDir = Join-Path $PSScriptRoot "data"
$RawDir  = Join-Path $DataDir "raw"
$LogsDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $DataDir,$RawDir,$LogsDir | Out-Null

function Write-Log($msg) {
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$ts] $msg"
  Write-Host $line
  Add-Content -Path (Join-Path $LogsDir ("update_" + (Get-Date).ToString("yyyyMMdd") + ".log")) -Value $line -Encoding UTF8
}

function Save-Status($ok, $message, $extra=@{}) {
  $status = [ordered]@{
    updated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    ok = [bool]$ok
    message = $message
  }
  foreach ($k in $extra.Keys) { $status[$k] = $extra[$k] }
  ($status | ConvertTo-Json -Depth 8) | Set-Content -Path (Join-Path $DataDir "status.json") -Encoding UTF8
}

# ------------------------------
# Best-effort race summary extraction (schema-agnostic)
# ------------------------------
function Find-RaceArray($node, [int]$depth = 0) {
  if ($null -eq $node) { return $null }
  if ($depth -gt 5) { return $null }

  # arrays
  if ($node -is [System.Collections.IEnumerable] -and -not ($node -is [string])) {
    $arr = @($node)
    if ($arr.Count -gt 0 -and ($arr[0] -is [pscustomobject] -or $arr[0] -is [hashtable])) {
      $keys = @($arr[0].PSObject.Properties.Name)
      if ($keys -match 'race' -or $keys -match 'race_no' -or $keys -match 'raceno' -or $keys -match 'rno') { return $arr }
      if ($keys -match 'payout' -or $keys -match 'sanrentan' -or $keys -match 'trifecta' -or $keys -match '3連単') { return $arr }
    }
  }

  # objects
  if ($node -is [pscustomobject] -or $node -is [hashtable]) {
    foreach ($p in $node.PSObject.Properties) {
      $found = Find-RaceArray $p.Value ($depth + 1)
      if ($null -ne $found) { return $found }
    }
  }
  return $null
}

function Try-GetInt($v) {
  if ($null -eq $v) { return $null }
  try {
    if ($v -is [int]) { return $v }
    if ($v -is [long]) { return [int]$v }
    $s = "$v".Trim()
    if ($s -match '^[0-9]+$') { return [int]$s }
    $s2 = ($s -replace '[^0-9]', '')
    if ($s2 -match '^[0-9]+$') { return [int]$s2 }
  } catch {}
  return $null
}

function Extract-RaceSummaries($json, [string]$date) {
  $races = Find-RaceArray $json
  if ($null -eq $races) { return @() }

  $out = @()
  foreach ($r in $races) {
    $props = $r.PSObject.Properties.Name

    $raceNo = $null
    foreach ($k in @('race_no','race','raceNo','raceno','rno','rnum')) {
      if ($props -contains $k) { $raceNo = Try-GetInt $r.$k; if ($raceNo) { break } }
    }

    $stadium = $null
    foreach ($k in @('stadium_code','stadium','venue','place','jyo','jyo_cd','stadiumCode')) {
      if ($props -contains $k) { $stadium = "$($r.$k)"; if ($stadium) { break } }
    }

    $payout = $null
    foreach ($k in @('sanrentan','trifecta','payout','payout_3t','payout3t','payout_trifecta','3t','3連単')) {
      if ($props -contains $k) { $payout = Try-GetInt $r.$k; if ($payout) { break } }
    }

    $out += [pscustomobject]@{
      date = $date
      stadium = $stadium
      race_no = $raceNo
      payout_3t = $payout
    }
  }
  return $out
}

function Notify-Fail($title, $body) {
  # Try Windows Toast notification (best effort). If it fails, fall back to msg.exe.
  try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null
    $template = [Windows.UI.Notifications.ToastTemplateType]::ToastText02
    $xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($template)
    $texts = $xml.GetElementsByTagName("text")
    $texts.Item(0).AppendChild($xml.CreateTextNode($title)) | Out-Null
    $texts.Item(1).AppendChild($xml.CreateTextNode($body)) | Out-Null
    $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
    $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("BPI")
    $notifier.Show($toast)
    return
  } catch {}
  try { & msg.exe * "$title`n$body" } catch {}
}

function Get-JstDate([int]$offsetDays) {
  # Convert "now" to JST and then shift days.
  $utc = [DateTimeOffset]::UtcNow
  $jst = $utc.ToOffset([TimeSpan]::FromHours(9))
  return $jst.Date.AddDays($offsetDays)
}

function Download-Json($url, $outPath) {
  Write-Log "GET $url"
  if ($DryRun) { return $null }
  $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60
  if ($resp.StatusCode -ne 200) { throw "HTTP $($resp.StatusCode)" }
  $resp.Content | Set-Content -Path $outPath -Encoding UTF8
  return $outPath
}

function First-Existing($obj, [string[]]$keys) {
  foreach ($k in $keys) {
    if ($null -ne $obj.PSObject.Properties[$k]) { return $obj.$k }
  }
  return $null
}

try {
  Write-Log "=== BPI updater start ==="

  # Seed (player master) is optional for now, but we validate the expected location.
  $seedDir = Join-Path $PSScriptRoot "seed"
  $playerXlsx = Join-Path $seedDir "BR_10Y_player_master_from_fan_files.xlsx"
  if (-not (Test-Path $seedDir)) {
    New-Item -ItemType Directory -Force -Path $seedDir | Out-Null
  }
  if (Test-Path $playerXlsx) {
    Write-Log "Found player master: $playerXlsx"
  } else {
    Write-Log "NOTE: player master not found yet (OK for motor-only build): $playerXlsx"
  }

  # === Backfill results (default: last 30 days) ===
  # We download by date (JST) and store as data/raw/YYYYMMDD_results.json
  # This makes the first run a 'one-shot' 30-day build. Later runs only add missing days.
  $latestTarget = Get-JstDate -1   # yesterday JST
  $latestYmd = $latestTarget.ToString("yyyyMMdd")
  $downloaded = 0
  $skipped = 0

  for ($i = 1; $i -le $BackfillDays; $i++) {
    $d = Get-JstDate (-$i)
    $ymd = $d.ToString("yyyyMMdd")
    $yyyy = $d.ToString("yyyy")
    $outFile = Join-Path $RawDir "${ymd}_results.json"
    if (Test-Path $outFile) {
      $skipped++
      continue
    }
    $url = "https://boatraceopenapi.github.io/results/v2/$yyyy/$ymd.json"
    try {
      Download-Json $url $outFile | Out-Null
      $downloaded++
    } catch {
      # Some days may have no races / file may not exist. We don't treat it as fatal.
      Write-Log "WARN: skip $ymd ($($_.Exception.Message))"
      if (Test-Path $outFile) { Remove-Item $outFile -Force }
    }
    if (-not $DryRun) { Start-Sleep -Milliseconds 400 }
  }
  Write-Log "Backfill done: downloaded=$downloaded skipped(existing)=$skipped"

  # Prune raw files older than KeepDays
  $cutoff = (Get-JstDate -$KeepDays)
  Get-ChildItem $RawDir -Filter "*_results.json" | ForEach-Object {
    if ($_.BaseName -match '^(\d{8})_results$') {
      $d = [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null)
      if ($d -lt $cutoff) {
        Write-Log "Prune $($_.Name)"
        if (-not $DryRun) { Remove-Item $_.FullName -Force }
      }
    }
  }

  # Load last MotorWindowDays worth of raws
  $motorCut = (Get-JstDate -$MotorWindowDays)
  $rawFiles = Get-ChildItem $RawDir -Filter "*_results.json" | Where-Object {
    $_.BaseName -match '^(\d{8})_results$' -and ([datetime]::ParseExact($Matches[1], "yyyyMMdd", $null) -ge $motorCut)
  } | Sort-Object Name

  $motorStats = @{}  # key: "stadium|motor"
  $raceCount = 0
  $raceSummaries = @()  # flat race summaries for viewer
  foreach ($f in $rawFiles) {
    $jsonText = Get-Content $f.FullName -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($jsonText)) { continue }
    $data = $jsonText | ConvertFrom-Json

    # For viewer (schema-agnostic): extract what we can from this day's JSON
    try {
      $raceSummaries += Extract-RaceSummaries $data $f.BaseName
    } catch {}

    # Data shape differs; handle array root vs object.
    $races = @()
    if ($data -is [System.Collections.IEnumerable]) { $races = $data }
    elseif ($data.races) { $races = $data.races }
    elseif ($data.data) { $races = $data.data }
    else { $races = @($data) }

    foreach ($race in $races) {
      $raceCount++
      $stadium = First-Existing $race @("race_stadium_number","stadium","stadium_number","place","place_no")
      $raceNo  = First-Existing $race @("race_number","raceNo","race_no","rno","race")
      $boats   = First-Existing $race @("boats","boat","racers","members","entries")
      if ($null -eq $boats) { continue }

      foreach ($b in $boats) {
        $motor = First-Existing $b @("motor_number","motor_no","motorNo","motor","motor_num")
        $rank  = First-Existing $b @("rank","finish","result_rank","arrival","place")
        $racer = First-Existing $b @("racer_number","racer_id","racerNo","registration_number","toban")
        if ($null -eq $stadium -or $null -eq $motor -or $null -eq $rank) { continue }

        $key = "$stadium|$motor"
        if (-not $motorStats.ContainsKey($key)) {
          $motorStats[$key] = [ordered]@{
            stadium = $stadium
            motor = $motor
            starts = 0
            top2 = 0
            top3 = 0
            sum_finish = 0
          }
        }
        $s = $motorStats[$key]
        $s.starts++
        $s.sum_finish += [int]$rank
        if ([int]$rank -le 2) { $s.top2++ }
        if ([int]$rank -le 3) { $s.top3++ }
      }
    }
  }

  $motorOut = @()
  foreach ($kv in $motorStats.GetEnumerator()) {
    $s = $kv.Value
    $starts = [double]$s.starts
    if ($starts -le 0) { continue }
    $motorOut += [ordered]@{
      stadium = $s.stadium
      motor = $s.motor
      starts = $s.starts
      top2_rate = [Math]::Round(($s.top2 / $starts) * 100, 2)
      top3_rate = [Math]::Round(($s.top3 / $starts) * 100, 2)
      avg_finish = [Math]::Round(($s.sum_finish / $starts), 3)
    }
  }

  $motorPath = Join-Path $DataDir "motor_30d.json"
  ($motorOut | Sort-Object stadium, motor | ConvertTo-Json -Depth 6) | Set-Content -Path $motorPath -Encoding UTF8
  Write-Log "Wrote motor_30d.json: $($motorOut.Count) motors (from $($rawFiles.Count) days, races scanned ~ $raceCount)"

  # Viewer helper files
  try {
    ($raceSummaries | ConvertTo-Json -Depth 6) | Set-Content -Path (Join-Path $DataDir 'races_flat.json') -Encoding UTF8

    $summary = @()
    foreach ($g in ($raceSummaries | Group-Object date)) {
      $payouts = @($g.Group | Where-Object { $_.payout_3t } | ForEach-Object { [int]$_.payout_3t })
      $avg = if ($payouts.Count -gt 0) { [math]::Round(($payouts | Measure-Object -Average).Average, 0) } else { $null }
      $max = if ($payouts.Count -gt 0) { ($payouts | Measure-Object -Maximum).Maximum } else { $null }
      $summary += [ordered]@{ date = $g.Name; races_found = $g.Count; payout_avg_3t = $avg; payout_max_3t = $max }
    }
    ($summary | Sort-Object date -Descending | ConvertTo-Json -Depth 6) | Set-Content -Path (Join-Path $DataDir 'summary.json') -Encoding UTF8
  } catch {
    # ignore
  }

  Save-Status $true "OK" @{ target_date = $latestYmd; raw_days = $rawFiles.Count; motors = $motorOut.Count; backfill_days = $BackfillDays; downloaded = $downloaded; races_flat = $raceSummaries.Count }
  Write-Log "=== BPI updater success ==="
  exit 0
}
catch {
  $msg = $_.Exception.Message
  Write-Log "ERROR: $msg"
  Save-Status $false $msg
  Notify-Fail "GI updater failed" $msg
  exit 1
}